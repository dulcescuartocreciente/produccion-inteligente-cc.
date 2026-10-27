"""Formula persistence: PostgreSQL (Supabase) when configured; otherwise local SQLite."""
import os
import sqlite3
from datetime import datetime
import numpy as np
import pandas as pd
import requests


def _secret(name):
    try:
        import streamlit as st
        return str(st.secrets.get(name, '')).strip() or os.getenv(name, '').strip()
    except (FileNotFoundError, KeyError):
        return os.getenv(name, '').strip()


def formula_database_url():
    # Mantener la firma utilizada por app.py. Nunca usar DATABASE_URL antigua.
    url = _secret('SUPABASE_URL').rstrip('/')
    key = _secret('SUPABASE_KEY')
    if bool(url) != bool(key):
        raise ValueError('Configurá SUPABASE_URL y SUPABASE_KEY juntas en Streamlit Secrets.')
    return (url, key) if url else None


class FormulaStore:
    def __init__(self, url, sqlite_path):
        self.credentials = url
        self.sqlite_path = sqlite_path
        self.is_remote = bool(url)
        if self.is_remote:
            self.base_url, self.key = url
            if not self.base_url.startswith('https://'):
                raise ValueError('SUPABASE_URL debe comenzar con https://')

    def _request(self, method, resource, *, params=None, payload=None, prefer=None):
        headers = {
            'apikey': self.key,
            'Authorization': 'Bearer ' + self.key,
            'Accept-Profile': 'produccion_cc',
            'Content-Profile': 'produccion_cc',
        }
        if prefer:
            headers['Prefer'] = prefer
        try:
            response = requests.request(
                method, self.base_url + '/rest/v1/' + resource,
                headers=headers, params=params, json=payload, timeout=25,
            )
            response.raise_for_status()
        except requests.RequestException as exc:
            detail = ''
            if getattr(exc, 'response', None) is not None:
                try:
                    body = exc.response.json()
                    detail = str(body.get('message') or body.get('hint') or body.get('code') or '')
                except ValueError:
                    detail = 'Error HTTP ' + str(exc.response.status_code)
            raise RuntimeError('Error de acceso a Supabase (esquema produccion_cc): ' + (detail or type(exc).__name__) + '. Verificá esquema expuesto, permisos y clave.') from None
        if not response.content:
            return None
        try:
            return response.json()
        except ValueError:
            return None

    def connect(self):
        if self.is_remote:
            raise RuntimeError('El acceso remoto usa la API de Supabase, no psycopg.')
        return sqlite3.connect(self.sqlite_path, timeout=20)

    def initialize(self):
        if self.is_remote:
            # Las tablas se crean UNA SOLA VEZ mediante el SQL de instalación.
            self._request('GET', 'formulas', params={'select': 'codigo', 'limit': '1'})
            self._request('GET', 'formula_ingredientes', params={'select': 'codigo', 'limit': '1'})
            return
        with self.connect() as con:
            con.execute('CREATE TABLE IF NOT EXISTS formulas (codigo TEXT, version INTEGER, producto TEXT, rendimiento_bejerman REAL, rendimiento_teorico REAL, observaciones TEXT, creado TEXT, origen TEXT, PRIMARY KEY(codigo,version))')
            con.execute('CREATE TABLE IF NOT EXISTS formula_ingredientes (codigo TEXT, version INTEGER, posicion INTEGER, grupo TEXT, ingrediente TEXT, cantidad REAL, unidad TEXT, brix REAL, PRIMARY KEY(codigo,version,posicion))')

    def query(self, sql, params=(), columns=None):
        if self.is_remote:
            raise RuntimeError('Las consultas SQL directas no se utilizan con Supabase REST.')
        with self.connect() as con:
            cur = con.execute(sql, params)
            rows = cur.fetchall()
            names = [x[0] for x in cur.description] if cur.description else columns or []
            return pd.DataFrame(rows, columns=names)

    def _all(self, table, order, params=None):
        # Paginación: PostgREST limita normalmente a 1.000 filas por respuesta.
        rows = []
        start = 0
        while True:
            q = {'select': '*', 'order': order, 'limit': 1000, 'offset': start}
            q.update(params or {})
            page = self._request('GET', table, params=q) or []
            rows.extend(page)
            if len(page) < 1000:
                return pd.DataFrame(rows)
            start += len(page)

    def versions(self):
        if self.is_remote:
            return self._all('formulas', 'producto.asc,version.desc')
        return self.query('SELECT * FROM formulas ORDER BY producto,version DESC')

    def ingredients(self, code, version):
        if self.is_remote:
            df = self._all('formula_ingredientes', 'posicion.asc', {'codigo': 'eq.' + str(code), 'version': 'eq.' + str(int(version))})
            return df.reindex(columns=['grupo', 'ingrediente', 'cantidad', 'unidad', 'brix'])
        return self.query('SELECT grupo,ingrediente,cantidad,unidad,brix FROM formula_ingredientes WHERE codigo=? AND version=? ORDER BY posicion', (code, int(version)))

    def all_ingredients(self):
        if self.is_remote:
            return self._all('formula_ingredientes', 'codigo.asc,version.asc,posicion.asc')
        return self.query('SELECT * FROM formula_ingredientes ORDER BY codigo,version,posicion')

    def save(self,code,product,bejerman,theoretical,notes,ingredients,origin):
        code=str(code).strip().upper();product=str(product).strip()
        if not code or not product:raise ValueError('Completá código y producto.')
        if bejerman is not None and (not np.isfinite(float(bejerman)) or float(bejerman)<=0):raise ValueError('El rendimiento Bejerman debe ser positivo cuando está informado.')
        rows=[]
        for _,r in ingredients.iterrows():
            name=str(r.get('ingrediente','')).strip()
            if not name or name.lower()=='nan':continue
            qty=pd.to_numeric(r.get('cantidad'),errors='coerce')
            if pd.isna(qty) or qty<0:raise ValueError('Cantidad inválida de ingrediente.')
            brix=pd.to_numeric(r.get('brix'),errors='coerce')
            rows.append((str(r.get('grupo','') or ''),name,float(qty),str(r.get('unidad','') or ''),None if pd.isna(brix) else float(brix)))
        if not rows:raise ValueError('La fórmula no contiene ingredientes válidos.')
        theoretical=None if pd.isna(theoretical) else float(theoretical)
        if self.is_remote:
            payload = {
                'p_codigo': code, 'p_producto': product,
                'p_bejerman': None if bejerman is None else float(bejerman), 'p_teorico': theoretical,
                'p_notas': str(notes or ''), 'p_origen': str(origin or ''),
                'p_ingredientes': [
                    {'grupo': r[0], 'ingrediente': r[1], 'cantidad': r[2], 'unidad': r[3], 'brix': r[4]}
                    for r in rows
                ],
            }
            return self._request('POST', 'rpc/guardar_formula', payload=payload)
        with self.connect() as con:
            q='?'
            cur=con.execute(f'SELECT version,producto,rendimiento_bejerman,rendimiento_teorico FROM formulas WHERE codigo={q} ORDER BY version DESC LIMIT 1',(code,))
            last=cur.fetchone()
            if last:
                old=con.execute(f'SELECT grupo,ingrediente,cantidad,unidad,brix FROM formula_ingredientes WHERE codigo={q} AND version={q} ORDER BY posicion',(code,last[0])).fetchall()
                def canonical(v):return '' if v is None else str(v).strip()
                same_items=len(old)==len(rows) and all(all(canonical(a)==canonical(b) for a,b in zip(oldrow,newrow)) for oldrow,newrow in zip(old,rows))
                old_theo=last[3]
                same_theo=(old_theo is None and theoretical is None) or (old_theo is not None and theoretical is not None and abs(float(old_theo)-theoretical)<1e-8)
                if last[1]==product and abs(float(last[2])-float(bejerman))<1e-8 and same_theo and same_items:
                    return None
            version=1 if last is None else int(last[0])+1
            sql='INSERT INTO formulas (codigo,version,producto,rendimiento_bejerman,rendimiento_teorico,observaciones,creado,origen) VALUES ('+','.join([q]*8)+')'
            con.execute(sql,(code,version,product,float(bejerman),theoretical,notes,datetime.now().isoformat(timespec='seconds'),origin))
            sql='INSERT INTO formula_ingredientes (codigo,version,posicion,grupo,ingrediente,cantidad,unidad,brix) VALUES ('+','.join([q]*8)+')'
            for i,row in enumerate(rows):con.execute(sql,(code,version,i,*row))
            return version
