"""Formula persistence: PostgreSQL (Supabase) when configured; otherwise local SQLite."""
import os
import sqlite3
from datetime import datetime
import numpy as np
import pandas as pd


def formula_database_url():
    try:
        import streamlit as st
        return str(st.secrets.get('DATABASE_URL', '')).strip() or os.getenv('DATABASE_URL','').strip()
    except (FileNotFoundError, KeyError):
        return os.getenv('DATABASE_URL','').strip()


class FormulaStore:
    def __init__(self, url, sqlite_path):
        self.url=url
        self.sqlite_path=sqlite_path
        self.is_remote=bool(url)

    def connect(self):
        if self.is_remote:
            import psycopg
            return psycopg.connect(self.url,connect_timeout=12)
        return sqlite3.connect(self.sqlite_path,timeout=20)

    def initialize(self):
        with self.connect() as con:
            if self.is_remote:
                con.execute('''CREATE TABLE IF NOT EXISTS formulas (codigo TEXT NOT NULL, version INTEGER NOT NULL, producto TEXT NOT NULL, rendimiento_bejerman DOUBLE PRECISION NOT NULL, rendimiento_teorico DOUBLE PRECISION, observaciones TEXT, creado TEXT NOT NULL, origen TEXT, PRIMARY KEY(codigo,version))''')
                con.execute('''CREATE TABLE IF NOT EXISTS formula_ingredientes (codigo TEXT NOT NULL, version INTEGER NOT NULL, posicion INTEGER NOT NULL, grupo TEXT, ingrediente TEXT NOT NULL, cantidad DOUBLE PRECISION, unidad TEXT, brix DOUBLE PRECISION, PRIMARY KEY(codigo,version,posicion), FOREIGN KEY(codigo,version) REFERENCES formulas(codigo,version))''')
            else:
                con.execute('CREATE TABLE IF NOT EXISTS formulas (codigo TEXT, version INTEGER, producto TEXT, rendimiento_bejerman REAL, rendimiento_teorico REAL, observaciones TEXT, creado TEXT, origen TEXT, PRIMARY KEY(codigo,version))')
                con.execute('CREATE TABLE IF NOT EXISTS formula_ingredientes (codigo TEXT, version INTEGER, posicion INTEGER, grupo TEXT, ingrediente TEXT, cantidad REAL, unidad TEXT, brix REAL, PRIMARY KEY(codigo,version,posicion))')

    def query(self, sql, params=(), columns=None):
        with self.connect() as con:
            cur=con.execute(sql.replace('?', '%s') if self.is_remote else sql,params)
            rows=cur.fetchall()
            names=[x[0] for x in cur.description] if cur.description else columns or []
            return pd.DataFrame(rows,columns=names)

    def versions(self):
        return self.query('SELECT * FROM formulas ORDER BY producto,version DESC')

    def ingredients(self,code,version):
        return self.query('SELECT grupo,ingrediente,cantidad,unidad,brix FROM formula_ingredientes WHERE codigo=? AND version=? ORDER BY posicion',(code,int(version)))

    def all_ingredients(self):
        return self.query('SELECT * FROM formula_ingredientes ORDER BY codigo,version,posicion')

    def save(self,code,product,bejerman,theoretical,notes,ingredients,origin):
        code=str(code).strip().upper();product=str(product).strip()
        if not code or not product:raise ValueError('Completá código y producto.')
        if not np.isfinite(float(bejerman)) or float(bejerman)<=0:raise ValueError('El rendimiento Bejerman debe ser positivo.')
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
        with self.connect() as con:
            if self.is_remote:
                # Serialize concurrent writes per product, preventing duplicate version numbers.
                con.execute('SELECT pg_advisory_xact_lock(hashtext(%s))',(code,))
            q='%s' if self.is_remote else '?'
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
