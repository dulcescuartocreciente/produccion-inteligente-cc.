import io
import calendar
import html
import hashlib
import sqlite3
import zipfile
import os
import re
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Producción Inteligente CC', page_icon='🏭', layout='wide')
st.markdown('''<style>
.block-container{padding-top:1rem;max-width:1480px;padding-bottom:3rem}
[data-testid="stMetric"]{background:#f1f8fd;border:1px solid #d8e9f6;border-radius:12px;padding:10px;min-width:0}
[data-testid="stMetricLabel"]{font-size:.79rem!important;color:#24445c}
[data-testid="stMetricValue"]{font-size:clamp(.94rem,1.35vw,1.3rem)!important;overflow-wrap:anywhere;line-height:1.25}
h1,h2,h3{color:#175d8b}
.stTabs [data-baseweb="tab-list"]{gap:.35rem;flex-wrap:wrap;background:#eef6fc;border-radius:12px;padding:.35rem}
.stTabs [data-baseweb="tab"]{border-radius:9px;padding:.55rem .9rem;font-weight:600}
.stTabs [aria-selected="true"]{background:white;box-shadow:0 1px 5px #c9ddec}
[data-testid="stImage"] img{object-fit:contain!important;max-height:105px;width:auto!important}
</style>''', unsafe_allow_html=True)
from zoneinfo import ZoneInfo
logo = Path(__file__).with_name('logo_cc.png')
head_logo, head_text = st.columns([1, 7], vertical_alignment='center')
with head_logo:
    if logo.exists():
        st.image(str(logo), width=165)
with head_text:
    st.title('Producción Inteligente CC')
    st.caption('Cuarto Creciente · Gestión de cobranzas, rentabilidad y producción')

@st.fragment(run_every='60s')
def show_local_clock():
    now=datetime.now(ZoneInfo('America/Argentina/Buenos_Aires'))
    st.caption(f'📅 {now:%d/%m/%Y}  ·  🕒 {now:%H:%M} hs (Argentina)')
show_local_clock()


def clean(x):
    x = unicodedata.normalize('NFKD', str(x or ''))
    return ''.join(c for c in x if not unicodedata.combining(c)).strip().upper()


def normalize_client(x):
    return ' '.join(clean(x).split())


def excel_date(series):
    values = []
    for v in series:
        if pd.isna(v):
            values.append(pd.NaT)
        elif isinstance(v, (datetime, pd.Timestamp)):
            values.append(pd.Timestamp(v).normalize())
        elif isinstance(v, (int, float, np.integer, np.floating)):
            values.append(pd.to_datetime(float(v),unit='D',origin='1899-12-30',errors='coerce'))
        else:
            values.append(pd.to_datetime(v,errors='coerce',dayfirst=True))
    return pd.Series(values,index=series.index,dtype='datetime64[ns]')


def parse_amount(value):
    if pd.isna(value):
        return np.nan
    if isinstance(value, (int, float)):
        return float(value)
    import re
    v = re.sub(r'[^0-9,.-]', '', str(value))
    if not v:
        return np.nan
    if ',' in v:
        v = v.replace('.', '').replace(',', '.')
    elif v.count('.') > 1:
        v = v.replace('.', '')
    try:
        return float(v)
    except ValueError:
        return np.nan


def payment_days(value):
    import re
    if pd.isna(value) or not str(value).strip():
        return np.nan, 'SIN CONDICIÓN'
    term = clean(value)
    if isinstance(value, (int, float)) and float(value).is_integer() and 0 <= value <= 365:
        return int(value), 'AUTOMÁTICA'
    if re.fullmatch(r'\d{1,3}(?:[.,]0+)?', term):
        return int(float(term.replace(',', '.'))), 'AUTOMÁTICA'
    match = re.fullmatch(r'(\d{1,3})\s*DIAS?', term)
    if match:
        return int(match.group(1)), 'AUTOMÁTICA'
    if term in ('CONTADO', 'EFECTIVO', 'PAGO CONTADO', 'ANTICIPADO'):
        return 0, 'AUTOMÁTICA'
    if term in ('RETIRAR PAGO', 'RETIRA PAGO') or 'E-CHEQ' in term or 'ECHEQ' in term:
        return np.nan, 'FECHA A CONFIRMAR'
    return np.nan, 'CONDICIÓN NO RECONOCIDA'




def read_historical(file):
    """Lee cobros asignados a fechas en PEDIDOS ENTREGADOS del Cash Flow.

    IMPORTE es el valor de venta, NO un cobro confirmado. Solo se proyectan
    importes positivos colocados bajo columnas de fecha de cobranza.
    """
    book = pd.ExcelFile(file)
    records = []
    sheets = [name for name in book.sheet_names if clean(name) == 'PEDIDOS ENTREGADOS']
    if not sheets:
        raise ValueError('El Cash Flow no contiene la pestaña PEDIDOS ENTREGADOS.')
    for sheet in sheets:
        raw = pd.read_excel(book, sheet_name=sheet, header=None, dtype=object)
        for i in range(min(30, len(raw))):
            labels = [clean(v) if pd.notna(v) else '' for v in raw.iloc[i].tolist()]
            if not all(x in labels for x in ('FECHA', 'CLIENTE', 'IMPORTE')):
                continue
            c, a, f = (labels.index(x) for x in ('CLIENTE', 'IMPORTE', 'FECHA'))
            dated_cols = []
            for j, value in enumerate(raw.iloc[i].tolist()):
                if j in (c, a, f) or pd.isna(value):
                    continue
                if isinstance(value, (int, float, np.integer, np.floating)) and 40000 <= float(value) <= 60000:
                    dt = pd.to_datetime(float(value), unit='D', origin='1899-12-30', errors='coerce')
                else:
                    dt = pd.to_datetime(value, errors='coerce', dayfirst=True)
                if pd.notna(dt) and pd.Timestamp('2020-01-01') <= dt <= pd.Timestamp('2045-12-31'):
                    dated_cols.append((j, pd.Timestamp(dt).normalize()))
            for k in range(i + 1, len(raw)):
                row = raw.iloc[k]
                client = row.iloc[c]
                if pd.isna(client) or not str(client).strip():
                    continue
                delivery = excel_date(pd.Series([row.iloc[f]])).iloc[0]
                sale_amount = parse_amount(row.iloc[a])
                if pd.isna(delivery) or pd.isna(sale_amount) or sale_amount <= 0:
                    continue
                for j, date in dated_cols:
                    payment = parse_amount(row.iloc[j])
                    if pd.notna(payment) and payment > 0:
                        records.append({'FECHA ENTREGA': delivery, 'CLIENTE': str(client).strip(),
                                        'IMPORTE VENTA': float(sale_amount), 'FECHA COBRANZA': date,
                                        'IMPORTE': float(payment), 'ORIGEN': 'CASH FLOW'})
            break
    columns = ['FECHA ENTREGA', 'CLIENTE', 'IMPORTE VENTA', 'FECHA COBRANZA', 'IMPORTE', 'ORIGEN']
    return pd.DataFrame(records, columns=columns)


def read_obligations(file):
    """Extrae obligaciones sin inventar fechas ni asumir que siguen impagas.

    OBLIGACIONES tiene tres bloques independientes: B/C, F/G e I/J.
    Son listados de referencia; algunos pueden superponerse.
    """
    book = pd.ExcelFile(file)
    sheet = next((n for n in book.sheet_names if clean(n) == 'OBLIGACIONES'), None)
    if sheet is None:
        return pd.DataFrame(columns=['ID','GRUPO','CONCEPTO','MONTO'])
    raw = pd.read_excel(book, sheet_name=sheet, header=None, dtype=object)
    result = []
    for group, desc_col, amount_col in [('ADMINISTRACIÓN',1,2),('COMPRAS',5,6),('OTRAS OBLIGACIONES',8,9)]:
        if raw.shape[1] <= amount_col:
            continue
        for i in range(2,len(raw)):
            name = raw.iat[i,desc_col]
            if pd.isna(name) or not str(name).strip():
                continue
            label = str(name).strip()
            if clean(label) in ('TOTAL','MONTO','SUELDOS','OBLIGACIONES','0') or clean(label).startswith('OBLIGACIONES '):
                continue
            amount = parse_amount(raw.iat[i,amount_col])
            if pd.isna(amount) or amount <= 0:
                continue
            result.append({'ID':f'{group} · fila {i+1}', 'GRUPO':group,'CONCEPTO':label,'MONTO':float(amount)})
    return pd.DataFrame(result,columns=['ID','GRUPO','CONCEPTO','MONTO'])


def build_forecast(routes, cashflow):
    """Cash Flow tiene prioridad cuando una misma venta también aparece en rutas."""
    parts = []
    excluded = 0
    if not cashflow.empty:
        parts.append(cashflow.copy())
    if not routes.empty:
        rt = routes.copy()
        rt['FECHA COBRANZA'] = rt['FECHA'] + pd.to_timedelta(rt['DÍAS DE PAGO'], unit='D')
        rt = rt[rt['DÍAS DE PAGO'].notna() & rt['IMPORTE'].notna() & (rt['IMPORTE'] > 0)].copy()
        rt['FECHA ENTREGA'] = rt['FECHA']
        rt['IMPORTE VENTA'] = rt['IMPORTE']
        rt['ORIGEN'] = 'HOJA DE RUTA'
        if not cashflow.empty:
            cf_keys = set(zip(cashflow['CLIENTE'].map(normalize_client),
                              pd.to_datetime(cashflow['FECHA ENTREGA']).dt.strftime('%Y-%m-%d'),
                              cashflow['IMPORTE VENTA'].round(2)))
            duplicate = [ (normalize_client(r.CLIENTE),r._asdict()['FECHA ENTREGA'].strftime('%Y-%m-%d'),round(r._asdict()['IMPORTE VENTA'],2)) in cf_keys for r in rt.itertuples(index=False, name=None)] if False else [
                (normalize_client(row['CLIENTE']), row['FECHA ENTREGA'].strftime('%Y-%m-%d'),round(float(row['IMPORTE VENTA']),2)) in cf_keys
                for _,row in rt.iterrows()]
            excluded = sum(duplicate)
            rt = rt.loc[~pd.Series(duplicate,index=rt.index)]
        parts.append(rt[['FECHA ENTREGA','CLIENTE','IMPORTE VENTA','FECHA COBRANZA','IMPORTE','ORIGEN']])
    if not parts:
        return pd.DataFrame(columns=['FECHA ENTREGA','CLIENTE','IMPORTE VENTA','FECHA COBRANZA','IMPORTE','ORIGEN']), excluded
    return pd.concat(parts,ignore_index=True),excluded


def audit_route_cashflow(routes, cashflow):
    """Cruce de control, NO conciliación automática: cliente+fecha no identifica un remito."""
    cols=['FECHA ENTREGA','CLIENTE','CONDICIÓN DE PAGO','ESTADO CONDICIÓN',
          'COINCIDENCIAS CASH FLOW','VENTAS CASH FLOW','ESTADO CRUCE']
    if routes.empty:
        return pd.DataFrame(columns=cols)
    rt=routes.copy()
    rt['FECHA ENTREGA']=pd.to_datetime(rt['FECHA'],errors='coerce').dt.normalize()
    rt['CLAVE CLIENTE']=rt['CLIENTE'].map(normalize_client)
    if cashflow.empty:
        rt['COINCIDENCIAS CASH FLOW']=0
        rt['VENTAS CASH FLOW']=0.0
    else:
        cf=cashflow.copy()
        cf['FECHA ENTREGA']=pd.to_datetime(cf['FECHA ENTREGA'],errors='coerce').dt.normalize()
        cf['CLAVE CLIENTE']=cf['CLIENTE'].map(normalize_client)
        # Un importe de venta puede aparecer en varias cuotas: contar la venta una sola vez
        sales=cf[['CLAVE CLIENTE','FECHA ENTREGA','IMPORTE VENTA']].drop_duplicates()
        grouped=sales.groupby(['CLAVE CLIENTE','FECHA ENTREGA'],as_index=False).agg(
            **{'COINCIDENCIAS CASH FLOW':('IMPORTE VENTA','size'),
               'VENTAS CASH FLOW':('IMPORTE VENTA','sum')})
        rt=rt.merge(grouped,on=['CLAVE CLIENTE','FECHA ENTREGA'],how='left')
        rt['COINCIDENCIAS CASH FLOW']=rt['COINCIDENCIAS CASH FLOW'].fillna(0).astype(int)
        rt['VENTAS CASH FLOW']=rt['VENTAS CASH FLOW'].fillna(0.0)
    rt['ESTADO CRUCE']=rt['COINCIDENCIAS CASH FLOW'].map(
        lambda n:'SIN COINCIDENCIA' if n==0 else ('POSIBLE COINCIDENCIA - REVISAR' if n==1 else 'MÚLTIPLES VENTAS - REVISAR'))
    return rt[cols]


def money(v):
    return "$ " + f"{v:,.0f}".replace(",", ".")


def date_from_sheet(name):
    """Interpreta nombres diarios D[M]MYYYY / DDMMYYYY sin invertir día y mes."""
    digits = re.sub(r'\D', '', str(name))
    if len(digits) < 5:
        return pd.NaT
    for year_len in (4, 2):
        if len(digits) < year_len + 3:
            continue
        year = int(digits[-year_len:])
        if year_len == 2:
            year += 2000
        if not 2020 <= year <= 2045:
            continue
        front = digits[:-year_len]
        # Priorizar DDMM cuando hay cuatro cifras (1010 -> 10/10).
        candidates = [(2, 2)] if len(front) == 4 else [(1, 2), (2, 1)] if len(front) == 3 else []
        for day_len, month_len in candidates:
            try:
                day = int(front[:day_len]); month = int(front[day_len:day_len+month_len])
                return pd.Timestamp(year, month, day)
            except (ValueError, TypeError):
                continue
    return pd.NaT


def read_route(file):
    book = pd.ExcelFile(file)
    candidates=[]
    for name in book.sheet_names:
        raw=pd.read_excel(book,sheet_name=name,header=None)
        header_count=0
        for i in range(min(len(raw),40)):
            vals=[clean(v) for v in raw.iloc[i].tolist()]
            if 'CLIENTE' in vals and any(v.startswith('COBRAR') for v in vals):
                header_count+=1
        if header_count: candidates.append((name,raw,header_count))
    if not candidates: raise ValueError('No encontré una hoja con CLIENTE y COBRAR (días).')
    # Una hoja diaria: si el libro trae varias rutas, el usuario elige cuál importar.
    return candidates


def extract_route(sheet_name, raw, fallback_date):
    route_date=date_from_sheet(sheet_name)
    if pd.isna(route_date): route_date=pd.Timestamp(fallback_date)
    records=[]; active=None
    for i in range(len(raw)):
        values=raw.iloc[i].tolist(); row=[clean(v) for v in values]
        c=next((j for j,v in enumerate(row) if v=='CLIENTE'),None)
        a=next((j for j,v in enumerate(row) if v=='IMPORTE'),None)
        pay=next((j for j,v in enumerate(row) if v.startswith('COBRAR')),None)
        if c is not None and pay is not None:
            b=next((j for j,v in enumerate(row) if v.startswith('BULTOS')),None)
            active=(c,a,pay,b);continue
        if len(row)>1 and 'VUELTA' in row[1]: active=None;continue
        if active is None: continue
        c,a,pay,b=active
        client=values[c]
        if pd.isna(client) or not str(client).strip():continue
        name=str(client).strip()
        if any(x in clean(name) for x in ('TOTAL','CANCELADO','ANULADO')):continue
        # Las entregas marcadas como no realizadas no generan una cobranza prevista.
        if any('NO SALE' in clean(v) or 'NO SALIO' in clean(v) for v in values if pd.notna(v)):continue
        amount=parse_amount(values[a]) if a is not None else np.nan; condition=values[pay]
        bultos=values[b] if b is not None else np.nan
        if pd.isna(amount) and pd.isna(bultos) and pd.isna(condition):continue
        days,state=payment_days(condition)
        records.append({'FECHA':route_date,'CLIENTE':name,'BULTOS':bultos,'IMPORTE':amount,
                        'CONDICIÓN DE PAGO':'' if pd.isna(condition) else str(condition).strip(),
                        'DÍAS DE PAGO':days,'ESTADO CONDICIÓN':state,'ORIGEN HOJA':sheet_name,'FILA ORIGEN':i+1})
    if not records:raise ValueError('No encontré entregas en la hoja seleccionada.')
    return pd.DataFrame(records)


def signature(df):
    keys=['FECHA','CLIENTE','IMPORTE','CONDICIÓN DE PAGO','ORIGEN HOJA','FILA ORIGEN']
    normalized=df[keys].copy().fillna('')
    normalized['FECHA']=normalized['FECHA'].astype(str)
    return hashlib.sha256(normalized.to_csv(index=False).encode('utf-8')).hexdigest()


def store_path():
    return Path(__file__).with_name('produccion_cc_historial.sqlite3')


def init_db():
    with sqlite3.connect(store_path()) as con:
        con.execute('CREATE TABLE IF NOT EXISTS rutas (firma TEXT PRIMARY KEY, fecha_carga TEXT, hoja TEXT, fecha_ruta TEXT, registros INTEGER)')
        con.execute('CREATE TABLE IF NOT EXISTS operaciones (firma TEXT, fila INTEGER, fecha TEXT, cliente TEXT, bultos REAL, importe REAL, condicion TEXT, dias REAL, estado TEXT, hoja TEXT, PRIMARY KEY(firma,fila))')


def save_route(df):
    sig=signature(df)
    with sqlite3.connect(store_path()) as con:
        exists=con.execute('SELECT 1 FROM rutas WHERE firma=?',(sig,)).fetchone()
        if exists:return False
        con.execute('INSERT INTO rutas VALUES (?,?,?,?,?)',(sig,datetime.now().isoformat(),str(df['ORIGEN HOJA'].iloc[0]),str(df['FECHA'].iloc[0].date()),len(df)))
        for j,r in enumerate(df.itertuples(index=False)):
            d=df.iloc[j]
            def num(x):return None if pd.isna(x) else float(x)
            con.execute('INSERT INTO operaciones VALUES (?,?,?,?,?,?,?,?,?,?)',(
                sig,j,str(d['FECHA'].date()),str(d['CLIENTE']),num(d['BULTOS']),num(d['IMPORTE']),
                str(d['CONDICIÓN DE PAGO']),num(d['DÍAS DE PAGO']),str(d['ESTADO CONDICIÓN']),str(d['ORIGEN HOJA'])))
    return True


def load_routes():
    with sqlite3.connect(store_path()) as con:
        df=pd.read_sql_query('SELECT * FROM operaciones ORDER BY fecha, firma, fila',con)
        batches=pd.read_sql_query('SELECT * FROM rutas ORDER BY fecha_ruta DESC',con)
    if not df.empty:
        df=df.rename(columns={'fecha':'FECHA','cliente':'CLIENTE','bultos':'BULTOS','importe':'IMPORTE','condicion':'CONDICIÓN DE PAGO','dias':'DÍAS DE PAGO','estado':'ESTADO CONDICIÓN','hoja':'ORIGEN HOJA'})
        df['FECHA']=pd.to_datetime(df['FECHA'])
    return df,batches


init_db()
with st.sidebar:
    st.header('Archivos del cálculo')
    route_file=st.file_uploader('Hoja de ruta del día',type=['xlsx','xls'],help='Subí la hoja de ruta correspondiente al día que querés incorporar. Debe tener CLIENTE y COBRAR (días). Si no trae IMPORTE, las filas quedan pendientes de vincular con Cash Flow. Si el libro contiene varios días, podrás elegir la pestaña.')
    history_file=st.file_uploader('Cash Flow / flujo de cobranzas (.xlsx)',type=['xlsx','xls'],key='cash_flow',help='Subí el Excel con las fechas e importes de cobranzas. Se sumará a la proyección evitando operaciones repetidas con las hojas de ruta.')
    st.caption('El mismo Cash Flow también puede contener obligaciones. Se mostrarán para revisión, sin descontarlas automáticamente.')
    profitability_file=st.file_uploader('Costos, márgenes y precios por bulto (.xlsx)',type=['xlsx','xls'],key='profitability',help='Subí el Excel con las columnas Artic.-Conc., COSTO2, MARGEN y precio. Los costos cero quedan excluidos de la simulación.')
    st.divider()
    st.header('Parámetros')
    initial_cash=st.number_input('Caja disponible inicial ($)',min_value=0.0,value=0.0,step=100000.0,help='Dinero que ya está disponible. No incluye ventas pendientes de cobro.')
    risk=st.slider('Margen de cobranza no realizada (%)',0,60,20,help='Porcentaje de cobranzas previstas que se descuenta por precaución.')
    horizon=st.selectbox('Horizonte de planificación',[7,15,30,60,90],index=2,format_func=lambda x:f'{x} días',help='Días a considerar desde la fecha inicial.')
    safety=st.number_input('Reserva de seguridad ($)',min_value=0.0,value=0.0,step=100000.0,help='Dinero que preferís no comprometer en producción.')
    reference_date=st.date_input('Fecha inicial de planificación',value=datetime.now().date(),format='DD/MM/YYYY',help='Fecha desde la cual se analizan las cobranzas proyectadas.')
    st.caption('Las cobranzas proyectadas no equivalen a dinero efectivamente recibido. Las obligaciones solo se descuentan si las seleccionás para reservarlas.')

main_fin, main_profit, main_formulas, main_stock, main_demand, main_plan = st.tabs([
    '💰 Cobranzas y Cash Flow', '📊 Rentabilidad', '🧪 Fórmulas y rendimientos', '📦 Stock e historial', '🚚 Demanda y distribución', '🏭 Plan de producción'
])

with main_fin:
    obligation_reserve = 0.0
    if history_file is not None:
        try:
            obligations = read_obligations(history_file)
            with st.expander('📋 Obligaciones del Cash Flow · revisión antes de reservar', expanded=False):
                st.caption('El Excel no tiene fechas de vencimiento estructuradas para estas partidas. Puede incluir obligaciones ya pagadas o repetidas entre bloques. Seleccioná solo las pendientes que querés reservar en este escenario; no se descuentan automáticamente.')
                if obligations.empty:
                    st.info('No se detectaron obligaciones con importes positivos en la pestaña OBLIGACIONES.')
                else:
                    st.dataframe(obligations[['GRUPO','CONCEPTO','MONTO']],hide_index=True,use_container_width=True)
                    options=obligations['ID'].tolist()
                    selected=st.multiselect('Obligaciones pendientes a reservar (selección manual)',options,default=[],key='obligations_selected',format_func=lambda x: f"{obligations.loc[obligations['ID']==x,'CONCEPTO'].iloc[0]} · {money(obligations.loc[obligations['ID']==x,'MONTO'].iloc[0])} · {x}")
                    obligation_reserve=float(obligations.loc[obligations['ID'].isin(selected),'MONTO'].sum())
                    st.metric('Reserva por obligaciones seleccionadas',money(obligation_reserve))
                    st.session_state['cc_obligation_reserve'] = obligation_reserve
                    st.caption('Se resta como reserva global desde el comienzo del horizonte, no como egreso con fecha. Evitá seleccionar importes que ya incluiste en la reserva de seguridad.')
        except Exception as exc:
            st.warning('No pude leer las obligaciones del Cash Flow: '+str(exc))
    st.subheader('Indicadores financieros')
    indicators=st.empty()
    with indicators.container():
        a,b,c,d=st.columns(4)
        for col,label in zip((a,b,c,d),('Ventas registradas','Cobranzas previstas','Cobranzas ajustadas','Techo preliminar')):
            col.metric(label,'—')

    if route_file:
        try:
            candidates=read_route(route_file)
            names=[x[0] for x in candidates]
            selected=st.selectbox('Día / pestaña a incorporar',names,index=names.index('9102026') if '9102026' in names else 0,help='Elegí la pestaña de la hoja de ruta que querés cargar. El sistema reconoce también la hoja 9102026.') if len(names)>1 else names[0]
            raw=next(x[1] for x in candidates if x[0]==selected)
            fallback=st.date_input('Fecha de entrega de esta hoja',value=datetime.now().date(),format='DD/MM/YYYY',help='Se usa si la fecha no puede identificarse a partir del nombre de la pestaña.')
            preview=extract_route(selected,raw,fallback)
            st.caption(f'Vista previa: {len(preview)} filas de la pestaña {selected}.')
            if preview['IMPORTE'].isna().any():
                st.info('Esta hoja de ruta no trae importes en todas las entregas. Se guardan sus condiciones de pago, pero no se inventan montos de cobranza.')
            st.dataframe(preview[['FECHA','CLIENTE','IMPORTE','CONDICIÓN DE PAGO','DÍAS DE PAGO','ESTADO CONDICIÓN']],hide_index=True,use_container_width=True)
            if st.button('Guardar hoja de ruta en el historial',help='Agrega esta hoja al historial local de la plataforma y evita guardar dos veces una carga idéntica.'):
                if save_route(preview):st.success('Hoja incorporada al historial.')
                else:st.info('Esta misma hoja ya estaba guardada. No se duplicó.')
        except Exception as e:
            st.error(f'No se pudo interpretar la hoja de ruta: {e}')

    st.warning('Almacenamiento provisional: el historial se guarda en la aplicación, pero Streamlit Cloud puede borrarlo al reiniciarse o volver a desplegarse. Descargá un respaldo después de cada carga. Para almacenamiento permanente y compartido conectaremos una base de datos externa.')

    try:
        saved,batches=load_routes()
    except Exception as e:
        st.error(f'No se pudo consultar el historial: {e}');st.stop()

    if not batches.empty:
        with st.expander('Hojas de ruta incorporadas',expanded=False):
            st.dataframe(batches.rename(columns={'fecha_carga':'CARGADA','hoja':'PESTAÑA','fecha_ruta':'FECHA RUTA','registros':'FILAS'}),hide_index=True,use_container_width=True)

    if not saved.empty:
        buff=io.BytesIO()
        with pd.ExcelWriter(buff,engine='xlsxwriter',datetime_format='dd/mm/yyyy') as writer:
            saved.to_excel(writer,sheet_name='Historial entregas',index=False)
            batches.to_excel(writer,sheet_name='Hojas cargadas',index=False)
        st.download_button('📥 Descargar respaldo del historial',buff.getvalue(),file_name='produccion_cc_respaldo_historial.xlsx',help='Guardá este archivo fuera de Streamlit para no perder los datos si se reinicia la aplicación.')

    if st.button('🔎 Analizar cobranzas', type='primary', use_container_width=True,
                 help='Proyecta los ingresos por fecha del Cash Flow y de las hojas de ruta, evitando operaciones repetidas.'):
        error = None
        cashflow = pd.DataFrame()
        if history_file is not None:
            try:
                cashflow = read_historical(history_file)
            except Exception as exc:
                error = str(exc)
        if error:
            st.error('No se pudo interpretar el Cash Flow: ' + error)
        else:
             # Las rutas sin importe no generan cobros ficticios: el monto viene del Cash Flow.
            route_without_amount = int(saved['IMPORTE'].isna().sum()) if not saved.empty else 0
            if route_without_amount:
                st.info(f'{route_without_amount} entrega(s) de hoja de ruta sin importe: sus montos se consultan en Cash Flow. No se agregan cobranzas por duplicado ni se inventan valores.')
            audit = audit_route_cashflow(saved, cashflow)
            st.session_state['route_cashflow_audit'] = audit
            forecast, excluded = build_forecast(saved, cashflow)
            st.session_state['analysis'] = (forecast, excluded, float(initial_cash), int(risk),
                                            int(horizon), float(safety), pd.Timestamp(reference_date))

    if 'route_cashflow_audit' in st.session_state:
        audit = st.session_state['route_cashflow_audit']
        if not audit.empty:
            st.subheader('Control de cruce: rutas y Cash Flow')
            st.caption('Coincidencias orientativas por cliente y fecha de entrega. NO se asignan importes a remitos sin identificador único ni se suman como cobros nuevos.')
            a,b,c=st.columns(3)
            a.metric('Entregas revisadas',len(audit))
            b.metric('Sin coincidencia',int((audit['COINCIDENCIAS CASH FLOW']==0).sum()))
            c.metric('Con varias ventas posibles',int((audit['COINCIDENCIAS CASH FLOW']>1).sum()))
            with st.expander('Ver entregas pendientes de conciliación'):
                st.dataframe(audit,hide_index=True,use_container_width=True)
                st.download_button('Descargar control de cruces (CSV)',audit.to_csv(index=False).encode('utf-8-sig'),file_name='control_cruce_rutas_cashflow.csv',mime='text/csv')

    if 'analysis' in st.session_state:
        forecast, excluded, cash, risk_used, horizon_used, reserve, start = st.session_state['analysis']
        if forecast.empty:
            st.warning('No encontramos cobranzas con fecha e importe. Cargá el Cash Flow o una hoja de ruta y presioná Analizar cobranzas.')
        else:
            forecast = forecast.copy()
            forecast['FECHA COBRANZA'] = pd.to_datetime(forecast['FECHA COBRANZA'])
            forecast['COBRANZA AJUSTADA'] = forecast['IMPORTE'] * (1-risk_used/100)
            end = start + pd.Timedelta(days=horizon_used)
            usable = forecast[forecast['FECHA COBRANZA'].between(start,end)].copy()
            expected = float(usable['IMPORTE'].sum())
            adjusted = float(usable['COBRANZA AJUSTADA'].sum())
            budget = max(0, cash + adjusted - reserve - obligation_reserve)
            with indicators.container():
                a,b,c,d = st.columns(4)
                a.metric('Caja inicial', money(cash))
                b.metric(f'Cobranzas ({horizon_used} días)', money(expected))
                c.metric('Ingresos prudentes', money(adjusted))
                d.metric('Techo preliminar', money(budget))
            st.caption('Simulación: cobranzas no confirmadas. Solo se descuentan las obligaciones seleccionadas manualmente; no se asignan fechas de vencimiento ni se presumen pagos realizados.')
            if excluded:
                st.info(f'Se excluyeron {excluded} operaciones de hojas de ruta porque ya estaban en el Cash Flow (mismo cliente, fecha de entrega e importe).')
            st.subheader('Calendario de ingresos para planificar producción')
            daily = usable.groupby('FECHA COBRANZA',as_index=False).agg(
                COBRANZA_PREVISTA=('IMPORTE','sum'), COBRANZA_AJUSTADA=('COBRANZA AJUSTADA','sum'))
            days = pd.DataFrame({'FECHA COBRANZA':pd.date_range(start,end,freq='D')})
            daily = days.merge(daily,on='FECHA COBRANZA',how='left').fillna(0)
            daily['CAJA PROYECTADA'] = cash + daily['COBRANZA_AJUSTADA'].cumsum()
            daily['TECHO PRELIMINAR'] = (daily['CAJA PROYECTADA']-reserve-obligation_reserve).clip(lower=0)
            st.markdown('#### 📅 Calendario mensual de cobranzas')
            st.caption('Importes previstos por fecha de cobro. El color identifica los días con ingresos proyectados; no significa que estén efectivamente cobrados.')
            calendar_months = pd.period_range(forecast['FECHA COBRANZA'].min().to_period('M'), forecast['FECHA COBRANZA'].max().to_period('M'), freq='M')
            if len(calendar_months) > 240:
                calendar_months = pd.period_range(start.to_period('M')-12, start.to_period('M')+12, freq='M')
            month_keys = [str(m) for m in calendar_months]
            default_month = str(start.to_period('M'))
            month_idx = month_keys.index(default_month) if default_month in month_keys else 0
            selected_month = st.selectbox('Mes a visualizar', month_keys, index=month_idx, key='collection_month', help='Elegí el mes de las cobranzas previstas. Podés consultar meses fuera del horizonte de planificación.')
            year, month = map(int, selected_month.split('-'))
            monthly = forecast.loc[(forecast['FECHA COBRANZA'].dt.year == year) & (forecast['FECHA COBRANZA'].dt.month == month)].copy()
            monthly['AJUSTADA'] = monthly['IMPORTE'] * (1-risk_used/100)
            monthly_by_day = monthly.groupby(monthly['FECHA COBRANZA'].dt.day).agg(prevista=('IMPORTE','sum'),ajustada=('AJUSTADA','sum'),operaciones=('IMPORTE','size')).to_dict('index')
            m1,m2,m3 = st.columns(3)
            m1.metric('Cobranza del mes', money(monthly['IMPORTE'].sum()))
            m2.metric('Cobranza prudente del mes', money(monthly['AJUSTADA'].sum()))
            m3.metric('Operaciones previstas', str(len(monthly)))
            weekdays = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom']
            cells = ['<div class="cc-weekday">'+x+'</div>' for x in weekdays]
            for week in calendar.monthcalendar(year,month):
                for day in week:
                    if day == 0:
                        cells.append('<div class="cc-day cc-empty"></div>')
                        continue
                    values = monthly_by_day.get(day)
                    css = ' cc-income' if values else ''
                    label = '<span class="cc-amount">'+html.escape(money(values['prevista']))+'</span><span class="cc-count">'+str(values['operaciones'])+' operación(es)</span>' if values else '<span class="cc-count">Sin cobranzas</span>'
                    cells.append('<div class="cc-day'+css+'"><span class="cc-number">'+str(day)+'</span>'+label+'</div>')
            st.markdown('''<style>
            .cc-month-grid{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:5px}
            .cc-weekday{text-align:center;font-weight:700;color:#285a77;padding:8px 0;background:#eaf4fa;border-radius:7px}
            .cc-day{border:1px solid #dce5eb;border-radius:9px;min-height:93px;padding:7px;display:flex;flex-direction:column;gap:6px;overflow-wrap:anywhere}
            .cc-income{background:#e7f4fd;border-color:#9bcde9}
            .cc-empty{border:none;background:transparent}
            .cc-number{font-size:.9rem;font-weight:700;color:#184a67}
            .cc-amount{font-weight:750;font-size:clamp(.65rem,.9vw,.92rem);color:#145b88}
            .cc-count{font-size:.69rem;color:#637d8b}
            @media(max-width:700px){.cc-day{min-height:78px;padding:4px}.cc-amount{font-size:.66rem}.cc-count{display:none}}
            </style>''' + '<div class="cc-month-grid">'+''.join(cells)+'</div>',unsafe_allow_html=True)
            with st.expander('Detalle de las cobranzas del mes'):
                st.dataframe(monthly.sort_values('FECHA COBRANZA'),hide_index=True,use_container_width=True)
            st.line_chart(daily.set_index('FECHA COBRANZA')[['CAJA PROYECTADA','TECHO PRELIMINAR']])
            t1,t2,t3 = st.tabs(['Ingresos por día','Detalle de cobranzas','Por origen'])
            with t1: st.dataframe(daily,hide_index=True,use_container_width=True)
            with t2: st.dataframe(usable.sort_values('FECHA COBRANZA'),hide_index=True,use_container_width=True)
            with t3:
                st.dataframe(usable.groupby('ORIGEN',as_index=False).agg(COBRANZA=('IMPORTE','sum'),OPERACIONES=('IMPORTE','size')),hide_index=True,use_container_width=True)
            output = io.BytesIO()
            with pd.ExcelWriter(output,engine='xlsxwriter',datetime_format='dd/mm/yyyy') as writer:
                daily.to_excel(writer,sheet_name='Calendario ingresos',index=False)
                forecast.to_excel(writer,sheet_name='Todas las cobranzas',index=False)
                pd.DataFrame({'CONCEPTO':['Caja inicial','Ingresos previstos','Ingresos prudentes','Reserva de seguridad','Obligaciones seleccionadas','Techo preliminar'],
                              'IMPORTE':[cash,expected,adjusted,reserve,obligation_reserve,budget]}).to_excel(writer,sheet_name='Resumen financiero',index=False)
            st.download_button('📥 Descargar proyección de ingresos',output.getvalue(),
                               file_name='produccion_cc_proyeccion_ingresos.xlsx',
                               help='Descarga las fechas, importes y caja acumulada para planificar producción.')
    else:
        st.info('Cargá el Cash Flow y/o guardá una hoja de ruta. Después presioná **Analizar cobranzas**.')



with main_profit:
    st.header('📊 Rentabilidad y planificación de producción')
    st.caption('Esta etapa cruza costos por bulto con el presupuesto preliminar de cobranzas. Todavía no considera pedidos pendientes, stock, egresos, capacidad productiva ni plazos de fabricación.')

    def read_profitability(upload):
        df = pd.read_excel(upload, sheet_name=0)
        colmap = {str(c).strip().lower(): c for c in df.columns}
        required = ['artic.-conc.', 'costo2', 'margen', 'precio']
        missing = [x for x in required if x not in colmap]
        if missing:
            raise ValueError('Faltan columnas: ' + ', '.join(missing) + '. Se esperan Artic.-Conc., COSTO2, MARGEN y precio.')
        result = df[[colmap[x] for x in required]].copy()
        result.columns = ['PRODUCTO', 'COSTO POR BULTO', 'MARGEN INFORMADO', 'PRECIO POR BULTO']
        result['PRODUCTO'] = result['PRODUCTO'].astype(str).str.strip()
        for col in ['COSTO POR BULTO', 'MARGEN INFORMADO', 'PRECIO POR BULTO']:
            result[col] = pd.to_numeric(result[col], errors='coerce')
        result = result[(result['PRODUCTO'] != '') & (result['PRODUCTO'].str.lower() != 'nan')].copy()
        result['GANANCIA POR BULTO'] = result['PRECIO POR BULTO'] - result['COSTO POR BULTO']
        result['MARGEN CALCULADO (%)'] = 100 * result['GANANCIA POR BULTO'] / result['PRECIO POR BULTO'].where(result['PRECIO POR BULTO'] > 0)
        result['ESTADO'] = 'Válido'
        result.loc[result['COSTO POR BULTO'].isna() | (result['COSTO POR BULTO'] <= 0), 'ESTADO'] = 'Revisar costo'
        result.loc[result['PRECIO POR BULTO'].isna() | (result['PRECIO POR BULTO'] <= 0), 'ESTADO'] = 'Revisar precio'
        result.loc[(result['ESTADO'] == 'Válido') & (result['GANANCIA POR BULTO'] <= 0), 'ESTADO'] = 'Sin margen positivo'
        return result

    if profitability_file is None:
        st.info('Subí el Excel de costos, márgenes y precios para comparar productos y simular la fabricación.')
    else:
        try:
            profitability = read_profitability(profitability_file)
            valid = profitability[profitability['ESTADO'] == 'Válido'].copy()
            p1,p2,p3 = st.columns(3)
            p1.metric('Productos registrados', str(len(profitability)))
            p2.metric('Productos con costo válido', str(len(valid)))
            p3.metric('Productos para revisar', str(len(profitability)-len(valid)))
            if len(valid) < len(profitability):
                st.warning('Los productos sin costo positivo, precio válido o margen positivo no se utilizan en la simulación. No se supone que su fabricación sea gratuita.')
            st.dataframe(profitability, use_container_width=True, hide_index=True,
                         column_config={'MARGEN INFORMADO': st.column_config.NumberColumn('Margen informado',format='%.2f'),
                                        'MARGEN CALCULADO (%)': st.column_config.NumberColumn('Margen calculado (%)',format='%.2f')})
            st.subheader('Simulación por producto')
            st.caption('Elegí un producto y una cantidad. Se calcula el capital necesario y la ganancia potencial si se venden todos los bultos. No es una orden de fabricación.')
            if not valid.empty:
                chosen = st.selectbox('Producto a simular', valid['PRODUCTO'].tolist(), key='sim_product', help='Solo aparecen productos con costo y precio válidos y ganancia positiva.')
                qty = st.number_input('Bultos a fabricar (simulación)', min_value=0, max_value=1000000, value=100, step=10,
                                      help='Cantidad hipotética. Cuando carguemos pedidos y stock, usaremos la necesidad real.')
                item = valid[valid['PRODUCTO'] == chosen].iloc[0]
                investment = float(qty) * float(item['COSTO POR BULTO'])
                potential_sales = float(qty) * float(item['PRECIO POR BULTO'])
                potential_profit = potential_sales - investment
                s1,s2,s3 = st.columns(3)
                s1.metric('Inversión requerida', money(investment))
                s2.metric('Venta potencial', money(potential_sales))
                s3.metric('Ganancia bruta potencial', money(potential_profit))
                if 'analysis' in st.session_state:
                    fc,exc,cc,rr,hh,res,start = st.session_state['analysis']
                    fc = fc.copy()
                    fc['FECHA COBRANZA'] = pd.to_datetime(fc['FECHA COBRANZA'])
                    horizon_end = start + pd.Timedelta(days=hh)
                    within = fc[fc['FECHA COBRANZA'].between(start,horizon_end)]
                    available = max(0.0, cc + float(within['IMPORTE'].sum()) * (1-rr/100) - res - obligation_reserve)
                    st.metric('Presupuesto preliminar según cobranzas', money(available))
                    if investment > available:
                        st.warning('Esta simulación supera el presupuesto preliminar de fabricación calculado con las cobranzas.')
                    else:
                        st.success('Esta simulación entra en el presupuesto preliminar, antes de descontar egresos y otras obligaciones.')
                    if float(item['COSTO POR BULTO']) > 0:
                        st.caption(f'Con ese presupuesto, el máximo teórico de este producto sería {int(available // float(item["COSTO POR BULTO"])):,} bultos, sin fabricar otros productos ni considerar demanda.'.replace(',', '.'))
                else:
                    st.info('Presioná «Analizar cobranzas» arriba para comparar esta inversión con el dinero proyectado.')
            output_profit = io.BytesIO()
            with pd.ExcelWriter(output_profit, engine='xlsxwriter') as writer:
                profitability.to_excel(writer, sheet_name='Rentabilidad por bulto', index=False)
            st.download_button('📥 Descargar análisis de rentabilidad', output_profit.getvalue(),
                               file_name='produccion_cc_rentabilidad.xlsx',
                               help='Exporta costos, precios, márgenes calculados y advertencias de validación.')
        except Exception as exc:
            st.error('No se pudo interpretar el Excel de rentabilidad: ' + str(exc))



with main_formulas:
    st.header('🧪 Fórmulas y rendimientos')
    st.caption('El plan utiliza el rendimiento Bejerman (incluye merma). El código Bejerman es opcional. Cada edición genera una versión nueva y conserva las anteriores.')

    from formula_store_cc import FormulaStore, formula_database_url
    formula_store = FormulaStore(formula_database_url(), store_path())

    def formulas_db():
        formula_store.initialize()

    def formula_versions():
        return formula_store.versions()

    def formula_ingredients(code,version):
        return formula_store.ingredients(code,version)

    def save_formula(code,product,bejerman,theoretical,notes,ingredients,origin):
        return formula_store.save(code,product,bejerman,theoretical,notes,ingredients,origin)

    def parse_formula_excel(upload, source_name=None):
        """Detecta encabezados y columnas por nombre, no por posición fija."""
        book = pd.ExcelFile(upload)
        origin = str(source_name or getattr(upload, 'name', '') or 'formula.xlsx')
        problems = []
        for sheet in book.sheet_names[:5]:
            raw = pd.read_excel(book, sheet_name=sheet, header=None, dtype=object)
            def val(r, c):
                return raw.iat[r, c] if 0 <= r < len(raw) and 0 <= c < len(raw.columns) else None
            def norm(v):
                return clean(v) if pd.notna(v) else ''
            def num(v):
                return parse_amount(v)
            name = None; code = None; revision = None
            bej = np.nan; theo = np.nan; items = []
            # Nombre: preferir título de producto; en su ausencia usar el nombre del archivo.
            for i in range(min(15, len(raw))):
                for j in range(min(5, len(raw.columns))):
                    v = val(i,j); t = norm(v)
                    if not t or len(t) > 110: continue
                    if any(k in t for k in ('MERMELADA','GELATINA','BEBIDA','JUGO','DULCE','INFUSION','POLVO','SUPLEMENTO','JALEA')) and not any(k in t for k in ('RENDIMIENTO','INGREDIENTE','FORMULA ORIGINAL','PROCEDIMIENTO','REVISION')):
                        name = str(v).strip(); break
                if name: break
            if not name: name = Path(origin).stem
            for i in range(min(15,len(raw))):
                labels = [norm(v) for v in raw.iloc[i].tolist()]
                for j,t in enumerate(labels):
                    if t in ('CODIGO','COD'):
                        c = val(i+1,j)
                        if pd.notna(c) and str(c).strip(): code = str(c).strip()
                    if t in ('VERSION','VERSION:','REVISION'):
                        c = val(i+1,j)
                        if pd.notna(c): revision = str(c).strip()
            for i in range(len(raw)):
                labels = [norm(v) for v in raw.iloc[i].tolist()]
                for j,t in enumerate(labels[:min(6,len(labels))]):
                    if 'RENDIMIENTO' not in t and 'RENDIMEINTO' not in t and 'RENDIMINETO' not in t: continue
                    if 'BEJERMAN' not in t and 'TEORICO' not in t: continue
                    candidates = [num(val(i,k)) for k in range(j+1,min(len(labels),j+9))]
                    candidates = [v for v in candidates if pd.notna(v) and v > 0]
                    if candidates:
                        if 'BEJERMAN' in t: bej = candidates[0]
                        elif 'TEORICO' in t: theo = candidates[0]
            for i in range(min(len(raw),90)):
                labels = [norm(v) for v in raw.iloc[i].tolist()]
                ingredient_col = next((j for j,t in enumerate(labels) if t in ('INGREDIENTE','INGREDIENTES','MATERIA PRIMA','MATERIAS PRIMAS')),None)
                qty_col = next((j for j,t in enumerate(labels) if t in ('CANTIDAD','CANTIDADES','CANT.','CANT','KG','PESO (KG)')),None)
                if ingredient_col is None or qty_col is None or ingredient_col == qty_col: continue
                unit_col = next((j for j,t in enumerate(labels) if t in ('UNID','UND','UNIDAD','UNIDADES','U.M.')),None)
                brix_col = next((j for j,t in enumerate(labels) if 'BRIX' in t and 'KG' not in t),None)
                group = ''; current = []; blanks = 0
                for r in range(i+1,min(len(raw),i+250)):
                    ing = val(r,ingredient_col); text = norm(ing)
                    qty = num(val(r,qty_col))
                    rowfirst = norm(val(r,0))
                    if any(x in rowfirst for x in ('MASA TOTAL','TOTAL MASA','DATOS FISICOQUIMICOS','RENDIMIENTO','RENDIMEINTO','OBSERVACIONES','REVISIONES')) or text in ('TOTAL','MASA TOTAL','SUBTOTAL'):
                        if current: break
                    if not text or text in ('INGREDIENTE','INGREDIENTES') or text.startswith('TOTAL '):
                        blanks += 1
                        if blanks > 7 and current: break
                        continue
                    blanks = 0
                    if pd.isna(qty) or qty < 0: continue
                    if ingredient_col != 0:
                        g = val(r,0)
                        if pd.notna(g) and str(g).strip(): group = str(g).strip()
                    unit = val(r,unit_col) if unit_col is not None else ''
                    brix = num(val(r,brix_col)) if brix_col is not None else np.nan
                    current.append({'grupo':group,'ingrediente':str(ing).strip(),
                                    'cantidad':float(qty),'unidad':str(unit).strip() if pd.notna(unit) else '',
                                    'brix':None if pd.isna(brix) else float(brix)})
                if len(current) > len(items): items = current
            if not items:
                problems.append(f'{sheet}: no se identificó una tabla de ingredientes con cantidades')
                continue
            # Si falta rendimiento, conservar ingredientes y marcar pendiente.
            if pd.isna(bej) or bej <= 0:
                bej = np.nan
            identifier = 'F-' + hashlib.sha256(origin.encode('utf-8')).hexdigest()[:16].upper()
            return {'codigo':identifier,'codigo_bejerman':code or '', 'producto':name,
                    'revision_original':revision,'bejerman':None if pd.isna(bej) else float(bej),
                    'teorico':float(theo) if pd.notna(theo) else 0.,
                    'ingredientes':pd.DataFrame(items)}
        raise ValueError('; '.join(problems[:3]) or 'No se encontró una hoja de fórmula interpretable')

    formulas_db()
    if formula_store.is_remote:
        st.success('Fórmulas conectadas a Supabase (esquema produccion_cc): guardado permanente activo.')
    else:
        st.warning('Fórmulas en SQLite local (provisional). Configurá SUPABASE_URL y SUPABASE_KEY en Streamlit Secrets para guardarlas permanentemente en Supabase.')
    f1,fzip,f2,f3=st.tabs(['📥 Importar fórmula','📦 Importar ZIP','📚 Fórmulas e historial','✏️ Editar y simular'])
    with f1:
        formula_file=st.file_uploader('Subir fórmula de elaboración (.xlsx)',type=['xlsx','xls'],key='formula_upload',help='Subí un Excel de fórmula como el modelo de Mermelada Stevia Arándano. Se leerán ingredientes, código y rendimiento Bejerman.')
        if formula_file:
            try:
                parsed=parse_formula_excel(formula_file, formula_file.name)
                st.write(f"**{parsed['producto']}** · Código Bejerman: {parsed['codigo_bejerman'] or 'Sin código'} · Revisión del Excel: {parsed['revision_original']}")
                x,y=st.columns(2)
                x.metric('Rendimiento Bejerman',f"{parsed['bejerman']:,.2f}" if parsed['bejerman'] is not None else 'Pendiente')
                y.metric('Rendimiento teórico (packs)',f"{parsed['teorico']:,.2f}")
                st.dataframe(parsed['ingredientes'],hide_index=True,use_container_width=True)
                import_note=st.text_input('Observaciones de importación',key='formula_import_note',help='Anotá el motivo de la carga o la revisión de la fórmula.')
                if st.button('Guardar fórmula como nueva versión',type='primary',key='formula_import_save'):
                    version=save_formula(parsed['codigo'],parsed['producto'],parsed['bejerman'],parsed['teorico'],import_note,parsed['ingredientes'],f'Excel: {formula_file.name}; código Bejerman {parsed["codigo_bejerman"] or "sin código"}; revisión origen {parsed["revision_original"]}')
                    if version is None: st.info('Esta fórmula es idéntica a la última versión guardada; no se duplicó.')
                    else: st.success(f'Fórmula guardada. Versión interna {version}. Se conservaron las versiones anteriores.')
            except Exception as exc:st.error('No se pudo importar la fórmula: '+str(exc))
    with fzip:
        st.subheader('Importación masiva de fórmulas desde ZIP')
        st.caption('Importa Excel con o sin código Bejerman. Se identifican por ruta dentro del ZIP; exige ingredientes, cantidades y rendimiento Bejerman positivo. Revisá la vista previa antes de guardar.')
        zip_upload=st.file_uploader('Subir ZIP de fórmulas',type=['zip'],key='formulas_zip')
        if zip_upload is not None:
            try:
                zip_bytes=zip_upload.getvalue()
                if len(zip_bytes)>80*1024*1024:
                    raise ValueError('El ZIP supera 80 MB. Dividilo en varios archivos.')
                parsed_files=[]; issues=[]; seen_codes=set()
                with zipfile.ZipFile(io.BytesIO(zip_bytes)) as archive:
                    entries=[x for x in archive.infolist() if not x.is_dir() and x.filename.lower().endswith(('.xlsx','.xls')) and not Path(x.filename).name.startswith(('~$','._'))]
                    if len(entries)>500:raise ValueError('El ZIP contiene más de 500 Excel. Dividilo en lotes.')
                    if not entries:raise ValueError('El ZIP no contiene fórmulas Excel.')
                    for entry in entries:
                        try:
                            if entry.file_size>20*1024*1024:raise ValueError('Excel mayor a 20 MB')
                            raw=archive.read(entry)
                            parsed=parse_formula_excel(io.BytesIO(raw), entry.filename)
                            code=parsed['codigo'].strip().upper()
                            if code in seen_codes:
                                raise ValueError(f'Código {code} repetido en este ZIP. Revisá qué versión querés cargar.')
                            seen_codes.add(code)
                            parsed_files.append((entry.filename,parsed))
                        except Exception as exc:
                            issues.append({'Archivo':entry.filename,'Problema':str(exc)})
                if issues:
                    st.warning(f'{len(issues)} archivo(s) no se pudieron interpretar. No se guardarán.')
                    st.dataframe(pd.DataFrame(issues),hide_index=True,use_container_width=True)
                if parsed_files:
                    preview_rows=[{'Archivo':filename,'Código Bejerman':data['codigo_bejerman'],'Producto':data['producto'],'Rendimiento Bejerman':data['bejerman'],'Estado':'Pendiente de rendimiento' if data['bejerman'] is None else 'Completa','Ingredientes':len(data['ingredientes'])} for filename,data in parsed_files]
                    st.dataframe(pd.DataFrame(preview_rows),hide_index=True,use_container_width=True)
                    chosen=st.multiselect('Seleccioná las fórmulas a incorporar',options=[name for name,_ in parsed_files],default=[name for name,_ in parsed_files],key='zip_selected')
                    note=st.text_input('Observaciones de la importación',key='zip_notes')
                    if st.button('Guardar fórmulas seleccionadas',type='primary',key='zip_commit'):
                        saved_count=0; skipped_count=0; errors=[]
                        for filename,data in parsed_files:
                            if filename not in chosen:continue
                            try:
                                version=save_formula(data['codigo'],data['producto'],data['bejerman'],data['teorico'],note,data['ingredientes'],f'ZIP: {zip_upload.name} / {filename}; código Bejerman {data["codigo_bejerman"] or "sin código"}')
                                if version is None:skipped_count+=1
                                else:saved_count+=1
                            except Exception as exc:errors.append({'Archivo':filename,'Problema':str(exc)})
                        st.success(f'{saved_count} fórmula(s) guardadas. {skipped_count} sin cambios, omitidas.')
                        if errors:st.dataframe(pd.DataFrame(errors),hide_index=True,use_container_width=True)
            except (zipfile.BadZipFile,ValueError,RuntimeError) as exc:
                st.error(f'No pude procesar el ZIP: {exc}')

    with f2:
        versions=formula_versions()
        if versions.empty:st.info('Todavía no hay fórmulas guardadas. Importá la primera desde Excel.')
        else:
            latest=versions.sort_values('version').drop_duplicates('codigo',keep='last')
            st.subheader('Fórmulas vigentes')
            latest['estado'] = latest['rendimiento_bejerman'].apply(lambda x: 'Pendiente de rendimiento' if pd.isna(x) or float(x)<=0 else 'Completa')
            st.dataframe(latest[['codigo','producto','version','estado','rendimiento_bejerman','rendimiento_teorico','creado']],hide_index=True,use_container_width=True)
            st.subheader('Historial de revisiones')
            st.dataframe(versions[['codigo','producto','version','rendimiento_bejerman','observaciones','creado','origen']],hide_index=True,use_container_width=True)
            export=io.BytesIO()
            with pd.ExcelWriter(export,engine='xlsxwriter') as writer:
                versions.to_excel(writer,sheet_name='Versiones',index=False)
                formula_store.all_ingredients().to_excel(writer,sheet_name='Ingredientes',index=False)
            st.download_button('📥 Descargar respaldo de fórmulas',export.getvalue(),file_name='respaldo_formulas_cc.xlsx',help='Guardá una copia externa: el almacenamiento local de Streamlit Cloud puede perderse al reiniciar.')
    with f3:
        versions=formula_versions()
        if versions.empty:st.info('Importá una fórmula para habilitar su edición y simulación.')
        else:
            options=versions.sort_values('version').drop_duplicates('codigo',keep='last')
            selected_code=st.selectbox('Producto / código',options['codigo'].tolist(),format_func=lambda c:f"{options.loc[options.codigo==c,'producto'].iloc[0]} ({c})",key='formula_code')
            relevant=versions[versions.codigo==selected_code]
            chosen_version=st.selectbox('Versión para consultar o editar',relevant.version.tolist(),key='formula_version')
            base=relevant[relevant.version==chosen_version].iloc[0]
            ingredients=formula_ingredients(selected_code,chosen_version)
            with st.form('formula_edit_form'):
                new_product=st.text_input('Nombre del producto',value=base['producto'])
                new_bej=st.number_input('Rendimiento Bejerman (packs por elaboración)',min_value=0.0,value=float(base['rendimiento_bejerman']) if pd.notna(base['rendimiento_bejerman']) else 0.0,format='%.4f',help='Rendimiento con merma. Es el utilizado para calcular elaboraciones completas.')
                new_theo=st.number_input('Rendimiento teórico (referencia)',min_value=0.0,value=float(base['rendimiento_teorico'] or 0),format='%.4f')
                edited=st.data_editor(ingredients,num_rows='dynamic',hide_index=True,use_container_width=True,key=f'edit_{selected_code}_{chosen_version}',column_config={'cantidad':st.column_config.NumberColumn('Cantidad',min_value=0),'ingrediente':st.column_config.TextColumn('Ingrediente',required=True)})
                edit_note=st.text_input('Motivo del cambio / observaciones',help='El historial conserva esta nota junto con la fecha de la nueva versión.')
                submitted=st.form_submit_button('Guardar cambios como nueva versión',type='primary')
            if submitted:
                try:
                    if not edit_note.strip():raise ValueError('Indicá el motivo de la modificación para conservar la trazabilidad.')
                    new_ver=save_formula(selected_code,new_product,new_bej,new_theo,edit_note,edited,f'Edición de versión {chosen_version}')
                    if new_ver is None: st.info('No hay cambios respecto de la última versión; no se duplicó.')
                    else: st.success(f'Nueva versión {new_ver} guardada. La versión {chosen_version} permanece en el historial.')
                except Exception as exc:st.error(str(exc))
            st.subheader('Simular necesidad de elaboraciones completas')
            demand=st.number_input('Bultos necesarios (después de descontar stock)',min_value=0,max_value=1000000,value=150,step=1,key='formula_demand')
            # Se descartan fracciones: no se pueden despachar packs incompletos.
            usable=int(np.floor(float(base['rendimiento_bejerman']))) if pd.notna(base['rendimiento_bejerman']) else 0
            if usable<1:st.warning('Pendiente de rendimiento Bejerman válido: esta fórmula no puede usarse para calcular elaboraciones.')
            else:
                batches=int(np.ceil(demand/usable))
                produced=batches*usable
                a,b,c=st.columns(3)
                a.metric('Elaboraciones completas',str(batches))
                b.metric('Bultos utilizables',str(produced))
                c.metric('Excedente estimado',str(produced-demand))
                st.caption(f'Se toman {usable} bultos completos por elaboración (rendimiento Bejerman {float(base["rendimiento_bejerman"]):.2f}). No se permiten elaboraciones parciales. Esta simulación no descuenta materias primas ni presupuesto.')

    if not formula_store.is_remote:
        st.warning('Las fórmulas siguen en SQLite provisional. Configurá Supabase antes de importar el ZIP definitivo. Los otros historiales de la app (stock, rutas y distribución) aún usan SQLite local.')
    else:
        st.info('Las fórmulas se guardan permanentemente en Supabase. Los historiales de stock, rutas y distribución siguen siendo provisionales en SQLite local; se migrarán en una etapa posterior.')



with main_stock:
    # ─── STOCK DE PRODUCTOS TERMINADOS: FOTOGRAFÍAS POR FECHA ───
    st.divider()
    st.header('📦 Stock e historial')
    st.caption('Stock expresado en bultos. Cada carga guarda una fotografía completa de las existencias; no se suman stocks de diferentes fechas.')

    def init_stock_db():
        with sqlite3.connect(store_path()) as con:
            con.execute('CREATE TABLE IF NOT EXISTS stock_cargas (id INTEGER PRIMARY KEY AUTOINCREMENT, fecha_stock TEXT NOT NULL, fecha_carga TEXT NOT NULL, archivo TEXT NOT NULL, firma TEXT UNIQUE NOT NULL, origen TEXT NOT NULL, observaciones TEXT)')
            con.execute('CREATE TABLE IF NOT EXISTS stock_detalle (carga_id INTEGER NOT NULL, codigo TEXT NOT NULL, descripcion TEXT NOT NULL, bultos REAL NOT NULL, PRIMARY KEY(carga_id,codigo), FOREIGN KEY(carga_id) REFERENCES stock_cargas(id))')

    def parse_stock_excel(upload):
        book=pd.ExcelFile(upload)
        for sheet in book.sheet_names:
            raw=pd.read_excel(book,sheet_name=sheet,header=None,dtype=object)
            for i in range(min(len(raw),35)):
                labels=[clean(v) for v in raw.iloc[i].tolist()]
                code_col=next((j for j,v in enumerate(labels) if v in ('COD','CODIGO','COD.','ARTICULO')),None)
                name_col=next((j for j,v in enumerate(labels) if v in ('DESCRIPCION','DESCRIPCIÓN','PRODUCTO','DETALLE')),None)
                qty_col=next((j for j,v in enumerate(labels) if v in ('STOCK','EXISTENCIA','EXISTENCIAS','CANTIDAD')),None)
                if None in (code_col,name_col,qty_col):continue
                data=raw.iloc[i+1:,[code_col,name_col,qty_col]].copy()
                data.columns=['codigo','descripcion','bultos']
                data['codigo']=data['codigo'].apply(lambda v: str(v).strip().upper() if pd.notna(v) else '')
                data['descripcion']=data['descripcion'].fillna('').astype(str).str.strip()
                data['bultos']=data['bultos'].apply(parse_amount)
                data=data[(data.codigo!='') & (data.descripcion!='') & data.bultos.notna()].copy()
                data=data[~data.codigo.str.contains(r'^(TOTAL|SUBTOTAL|COD)$',case=False,regex=True)]
                if data.empty:continue
                if data.codigo.duplicated().any():
                    repeated=', '.join(data.loc[data.codigo.duplicated(),'codigo'].head(5))
                    raise ValueError(f'El archivo tiene códigos repetidos ({repeated}). Revisá el reporte antes de guardarlo.')
                if (data.bultos<0).any():
                    st.warning('El archivo contiene stock negativo. Se conservará tal como figura en Bejerman y se marcará para revisión.')
                return data.sort_values('codigo').reset_index(drop=True),sheet
        raise ValueError('No encontré las columnas COD, DESCRIPCION y STOCK. Revisá que el archivo sea el reporte de stock de Bejerman.')

    def stock_snapshots():
        with sqlite3.connect(store_path()) as con:
            return pd.read_sql_query('SELECT * FROM stock_cargas ORDER BY fecha_stock DESC, id DESC',con)

    def stock_details(snapshot_id):
        with sqlite3.connect(store_path()) as con:
            return pd.read_sql_query('SELECT codigo,descripcion,bultos FROM stock_detalle WHERE carga_id=? ORDER BY codigo',con,params=(int(snapshot_id),))

    def save_stock(data,date_value,filename,origin='Excel',notes=''):
        normalized=data.sort_values('codigo').copy()
        payload=f"{date_value.isoformat()}\n"+normalized[['codigo','descripcion','bultos']].to_csv(index=False,float_format='%.6f')
        fingerprint=hashlib.sha256(payload.encode('utf-8')).hexdigest()
        with sqlite3.connect(store_path()) as con:
            if con.execute('SELECT 1 FROM stock_cargas WHERE firma=?',(fingerprint,)).fetchone():return None
            cur=con.execute('INSERT INTO stock_cargas (fecha_stock,fecha_carga,archivo,firma,origen,observaciones) VALUES (?,?,?,?,?,?)',(date_value.isoformat(),datetime.now().isoformat(timespec='seconds'),filename,fingerprint,origin,notes))
            con.executemany('INSERT INTO stock_detalle (carga_id,codigo,descripcion,bultos) VALUES (?,?,?,?)',[(cur.lastrowid,str(r.codigo),str(r.descripcion),float(r.bultos)) for r in normalized.itertuples(index=False)])
            return cur.lastrowid

    init_stock_db()
    s1,s2,s3,s4=st.tabs(['📥 Cargar stock','📦 Stock actual','📊 Comparar fechas','📚 Historial y respaldo'])
    with s1:
        st.info('Subí el reporte de productos terminados de Bejerman en formato .xlsx. La fecha representa cuándo se midió el stock, no cuándo se subió el archivo.')
        stock_file=st.file_uploader('Excel de stock (.xlsx)',type=['xlsx'],key='stock_upload',help='Reporte de Bejerman con COD, DESCRIPCION y STOCK. Las cantidades están expresadas en bultos.')
        stock_date=st.date_input('Fecha correspondiente al stock',value=datetime(2026,10,9).date(),key='stock_date',help='Elegí la fecha del reporte. Cada fecha puede tener distintas versiones si se corrige el stock.')
        stock_notes=st.text_input('Observaciones (opcional)',key='stock_notes',help='Por ejemplo: cierre del día, recuento físico o ajuste del reporte.')
        if stock_file is not None:
            try:
                preview,sheet=parse_stock_excel(stock_file)
                c1,c2,c3=st.columns(3)
                c1.metric('Productos',len(preview))
                c2.metric('Con stock positivo',int((preview.bultos>0).sum()))
                c3.metric('Bultos netos',f'{preview.bultos.sum():,.0f}'.replace(',','.'))
                st.dataframe(preview.rename(columns={'codigo':'Código','descripcion':'Producto','bultos':'Stock (bultos)'}),hide_index=True,use_container_width=True)
                if st.button('Guardar fotografía de stock',type='primary',key='stock_save',help='Guarda una nueva fotografía completa. Si ya existe exactamente la misma carga para esa fecha, no se duplica.'):
                    result=save_stock(preview,stock_date,stock_file.name,notes=stock_notes)
                    if result is None:st.info('Esta fotografía de stock ya estaba guardada para esa fecha. No se duplicó.')
                    else:st.success(f'Stock del {stock_date.strftime("%d/%m/%Y")} guardado como registro #{result}.')
            except Exception as exc:st.error(f'No pude interpretar el Excel de stock: {exc}')
    with s2:
        snaps=stock_snapshots()
        if snaps.empty:st.info('Todavía no hay stock guardado. Cargá el primer Excel en la pestaña anterior.')
        else:
            latest=snaps.iloc[0]
            current=stock_details(latest.id)
            st.caption(f'Última fotografía: {pd.to_datetime(latest.fecha_stock).strftime("%d/%m/%Y")} · registro #{int(latest.id)} · {latest.archivo}')
            a,b,c=st.columns(3)
            a.metric('Productos',len(current))
            b.metric('Con stock',int((current.bultos>0).sum()))
            c.metric('Bultos netos',f'{current.bultos.sum():,.0f}'.replace(',','.'))
            query=st.text_input('Buscar por código o producto',key='stock_search')
            if query:
                current=current[current.codigo.str.contains(query,case=False,regex=False)|current.descripcion.str.contains(query,case=False,regex=False)]
            st.dataframe(current.rename(columns={'codigo':'Código','descripcion':'Producto','bultos':'Stock (bultos)'}),hide_index=True,use_container_width=True)
            st.caption('El stock más reciente se selecciona por fecha del reporte, no por fecha de carga. Si existen varias versiones del mismo día, se utiliza la última guardada.')
    with s3:
        snaps=stock_snapshots()
        if len(snaps)<2:st.info('Cuando guardes al menos dos fotografías, podrás comparar los cambios de stock por producto.')
        else:
            labels={int(r.id):f'{pd.to_datetime(r.fecha_stock).strftime("%d/%m/%Y")} · registro #{int(r.id)}' for r in snaps.itertuples(index=False)}
            options=list(labels)
            a,b=st.columns(2)
            older=a.selectbox('Fotografía anterior',options,index=1,format_func=lambda k:labels[k],key='stock_compare_old')
            newer=b.selectbox('Fotografía posterior',options,index=0,format_func=lambda k:labels[k],key='stock_compare_new')
            if older==newer:st.warning('Elegí dos fotografías diferentes.')
            else:
                old=stock_details(older).rename(columns={'bultos':'Stock anterior','descripcion':'Descripción anterior'})
                new=stock_details(newer).rename(columns={'bultos':'Stock nuevo'})
                comparison=old.merge(new,on='codigo',how='outer')
                comparison['descripcion']=comparison['descripcion'].fillna(comparison['Descripción anterior'])
                comparison['Stock anterior']=comparison['Stock anterior'].fillna(0)
                comparison['Stock nuevo']=comparison['Stock nuevo'].fillna(0)
                comparison['Diferencia']=comparison['Stock nuevo']-comparison['Stock anterior']
                comparison=comparison[['codigo','descripcion','Stock anterior','Stock nuevo','Diferencia']].sort_values('Diferencia')
                st.metric('Variación neta de bultos',f'{comparison.Diferencia.sum():+,.0f}'.replace(',','.'))
                st.dataframe(comparison.rename(columns={'codigo':'Código','descripcion':'Producto'}),hide_index=True,use_container_width=True)
                st.caption('La diferencia entre fechas no representa necesariamente ventas: también puede incluir producción, entregas, devoluciones y ajustes.')
    with s4:
        snaps=stock_snapshots()
        if snaps.empty:st.info('No hay registros guardados todavía.')
        else:
            st.dataframe(snaps[['id','fecha_stock','fecha_carga','archivo','origen','observaciones']].rename(columns={'id':'Registro','fecha_stock':'Fecha stock','fecha_carga':'Fecha carga','archivo':'Archivo','origen':'Origen','observaciones':'Observaciones'}),hide_index=True,use_container_width=True)
            export=io.BytesIO()
            with sqlite3.connect(store_path()) as con:
                all_details=pd.read_sql_query('SELECT c.id AS registro,c.fecha_stock,d.codigo,d.descripcion,d.bultos FROM stock_cargas c JOIN stock_detalle d ON d.carga_id=c.id ORDER BY c.fecha_stock,c.id,d.codigo',con)
            with pd.ExcelWriter(export,engine='xlsxwriter') as writer:
                snaps.to_excel(writer,sheet_name='Cargas',index=False)
                all_details.to_excel(writer,sheet_name='Stock historico',index=False)
            st.download_button('📥 Descargar historial de stock (Excel)',export.getvalue(),file_name='historial_stock_cc.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',help='Guardá este respaldo fuera de Streamlit: el historial local no es permanente.')
            st.caption('Cada carga es una fotografía completa. Las versiones anteriores se conservan; no se reemplazan.')
    st.warning('⚠️ El historial de stock se guarda por ahora en SQLite local; las fórmulas usan Supabase si está configurado. Streamlit Cloud puede borrar esa información al reiniciar o redesplegar. Descargá respaldos y conectemos almacenamiento permanente antes de usarlo operativamente.')



# Distribución realizada: pedidos asignados ya entregados; pendientes son demanda no cubierta.
def read_distribution(upload):
    book = pd.ExcelFile(upload)
    needed = {'Pendientes','Stock restante'}
    missing = needed.difference(book.sheet_names)
    if missing:
        raise ValueError('Faltan pestañas del archivo de Distribución Inteligente: '+', '.join(sorted(missing)))
    pending = pd.read_excel(book,sheet_name='Pendientes')
    stock = pd.read_excel(book,sheet_name='Stock restante')
    pending.columns = [str(c).strip().upper() for c in pending.columns]
    stock.columns = [str(c).strip().upper() for c in stock.columns]
    for df,required in [(pending,['CODIGO','DESCRIPCION','PENDIENTE']),(stock,['CODIGO','DESCRIPCION','STOCK RESTANTE'])]:
        missing_cols = [c for c in required if c not in df.columns]
        if missing_cols: raise ValueError('Faltan columnas: '+', '.join(missing_cols))
        df['CODIGO'] = df['CODIGO'].astype(str).str.strip().str.upper()
        df.drop(df[~df['CODIGO'].str.match(r'^[A-Z0-9_-]+$')].index,inplace=True)
    pending['PENDIENTE'] = pd.to_numeric(pending['PENDIENTE'],errors='coerce').fillna(0).clip(lower=0)
    stock['STOCK RESTANTE'] = pd.to_numeric(stock['STOCK RESTANTE'],errors='coerce').fillna(0).clip(lower=0)
    pending = pending[pending['PENDIENTE']>0].copy()
    return pending,stock


def distribution_plan(pending,stock):
    demand = pending.groupby('CODIGO',as_index=False).agg(DESCRIPCION=('DESCRIPCION','first'),PENDIENTE=('PENDIENTE','sum'))
    available = stock.groupby('CODIGO',as_index=False)['STOCK RESTANTE'].sum()
    plan = demand.merge(available,on='CODIGO',how='left')
    plan['STOCK RESTANTE'] = plan['STOCK RESTANTE'].fillna(0)
    plan['NECESIDAD NETA'] = (plan['PENDIENTE']-plan['STOCK RESTANTE']).clip(lower=0)
    return plan.sort_values(['NECESIDAD NETA','PENDIENTE'],ascending=False)


with main_demand:
    st.header('🚚 Demanda y distribución')
    st.caption('Importá el Excel exportado por Distribución Inteligente CC después de que se haya realizado la distribución. No se vuelven a descontar los bultos entregados.')
    distribution_file = st.file_uploader('Distribución completa (.xlsx)',type=['xlsx'],key='distribution_file',help='Archivo exportado por Distribución Inteligente, con las pestañas Pendientes y Stock restante.')
    distribution_date = st.date_input('Fecha de distribución realizada',value=datetime.now(ZoneInfo('America/Argentina/Buenos_Aires')).date(),format='DD/MM/YYYY',key='distribution_date',help='Fecha efectiva de la distribución, no la fecha de carga del archivo.')
    if distribution_file is not None:
        try:
            pending, remaining = read_distribution(distribution_file)
            plan = distribution_plan(pending,remaining)
            st.info('Se toma el stock restante del archivo como stock posterior a esa distribución. Si existe un inventario físico más reciente, ese debe prevalecer para el plan definitivo.')
            a,b,c,d=st.columns(4)
            a.metric('Bultos pendientes',f"{pending['PENDIENTE'].sum():,.0f}")
            b.metric('Stock posterior a distribución',f"{remaining['STOCK RESTANTE'].sum():,.0f}")
            c.metric('Necesidad neta',f"{plan['NECESIDAD NETA'].sum():,.0f}")
            d.metric('Productos con faltante',str((plan['NECESIDAD NETA']>0).sum()))
            p1,p2,p3=st.tabs(['📋 Necesidad por producto','👥 Pendientes por cliente','📦 Stock posterior'])
            with p1: st.dataframe(plan,hide_index=True,use_container_width=True)
            with p2: st.dataframe(pending,hide_index=True,use_container_width=True)
            with p3: st.dataframe(remaining,hide_index=True,use_container_width=True)
            st.caption('La necesidad neta se calcula por código: máximo entre 0 y pendientes menos stock restante. Aún no equivale a una orden de fabricación: falta cruzarla con fórmulas, costos y caja.')
            if st.button('Guardar distribución en historial',key='save_distribution',help='Guarda la fotografía de pendientes y stock posterior a la distribución, sin repetir una carga idéntica.'):
                payload = distribution_file.getvalue()
                digest = hashlib.sha256(payload).hexdigest()
                with sqlite3.connect(store_path()) as conn:
                    conn.execute('CREATE TABLE IF NOT EXISTS distribuciones (hash TEXT PRIMARY KEY, fecha TEXT NOT NULL, nombre TEXT NOT NULL, fecha_carga TEXT NOT NULL, pendientes TEXT NOT NULL, stock TEXT NOT NULL)')
                    cursor = conn.execute('INSERT OR IGNORE INTO distribuciones VALUES (?,?,?,?,?,?)',(digest,str(distribution_date),distribution_file.name,datetime.now().isoformat(),pending.to_json(orient='records',force_ascii=False),remaining.to_json(orient='records',force_ascii=False)))
                    conn.commit()
                    inserted=cursor.rowcount>0
                if inserted: st.success('Distribución guardada en el historial local.')
                else: st.info('Esta distribución ya estaba guardada. No se duplicó.')
            output=io.BytesIO()
            with pd.ExcelWriter(output,engine='xlsxwriter') as writer:
                plan.to_excel(writer,sheet_name='Necesidad neta',index=False)
                pending.to_excel(writer,sheet_name='Pendientes',index=False)
                remaining.to_excel(writer,sheet_name='Stock posterior',index=False)
            st.download_button('📥 Descargar análisis de demanda',output.getvalue(),file_name='produccion_cc_demanda_distribucion.xlsx')
        except Exception as exc:
            st.error('No se pudo interpretar el archivo de distribución: '+str(exc))
    with sqlite3.connect(store_path()) as conn:
        conn.execute('CREATE TABLE IF NOT EXISTS distribuciones (hash TEXT PRIMARY KEY, fecha TEXT NOT NULL, nombre TEXT NOT NULL, fecha_carga TEXT NOT NULL, pendientes TEXT NOT NULL, stock TEXT NOT NULL)')
        stored = pd.read_sql_query('SELECT fecha AS FECHA, nombre AS ARCHIVO, fecha_carga AS CARGADO FROM distribuciones ORDER BY fecha DESC, fecha_carga DESC',conn)
    if not stored.empty:
        with st.expander('📚 Distribuciones guardadas (almacenamiento provisional)'):
            st.dataframe(stored,hide_index=True,use_container_width=True)
    st.warning('Los registros guardados localmente pueden perderse al reiniciar Streamlit Cloud. Para un historial definitivo falta conectar una base de datos persistente.')


# ─── V11 · PLAN DE PRODUCCIÓN ORIENTADO A RENTABILIDAD ───
with main_plan:
    from optimizer_cc import optimize_batches
    st.header('🏭 Plan de producción · máxima ganancia')
    st.caption('Optimiza elaboraciones completas con demanda pendiente real, rendimiento Bejerman, costo por bulto y presupuesto. No supone que los excedentes se venden.')
    st.info('Para empezar: cargá una distribución en «Demanda y distribución», costos en la barra lateral y las fórmulas en «Fórmulas y rendimientos».')

    # Distribución subida en esta sesión; se permite recuperar la última carga guardada.
    source_pending = source_stock = None
    source_date = None
    try:
        if distribution_file is not None:
            source_pending, source_stock = read_distribution(distribution_file)
            source_date = str(distribution_date)
        else:
            with sqlite3.connect(store_path()) as con:
                latest_distribution = con.execute('SELECT fecha,pendientes,stock FROM distribuciones ORDER BY fecha DESC, fecha_carga DESC LIMIT 1').fetchone()
            if latest_distribution:
                source_date=latest_distribution[0]
                source_pending=pd.read_json(io.StringIO(latest_distribution[1]),orient='records')
                source_stock=pd.read_json(io.StringIO(latest_distribution[2]),orient='records')
    except Exception as exc:
        st.error(f'No pude leer la distribución: {exc}')

    if source_pending is None or source_stock is None:
        st.warning('Falta importar o guardar una distribución realizada. Sin pedidos pendientes no podemos calcular rentabilidad respaldada por demanda.')
    elif profitability_file is None:
        st.warning('Falta subir el Excel de costos y precios por bulto en la barra lateral.')
    else:
        try:
            needs=distribution_plan(source_pending,source_stock)
            needs=needs[needs['NECESIDAD NETA']>0].copy()
            prices=read_profitability(profitability_file)
            versions=formula_versions()
            if needs.empty:
                st.success('No hay faltantes netos en la distribución seleccionada.')
            elif versions.empty:
                st.warning('Falta guardar fórmulas con rendimientos Bejerman. Importalas en su pestaña.')
            else:
                st.caption(f'Distribución de referencia: {source_date}. Si existe stock físico posterior, cargalo en «Stock e historial» y seleccioná su uso aquí.')
                source_stock_date=pd.Timestamp(source_date).date()
                snaps=stock_snapshots()
                recent=snaps[pd.to_datetime(snaps.fecha_stock).dt.date>source_stock_date] if not snaps.empty else pd.DataFrame()
                stock_option=st.radio('Stock para calcular faltantes',
                    ['Stock posterior de la distribución']+(['Última fotografía Bejerman posterior a la distribución'] if not recent.empty else []),
                    horizontal=True,key='v11_stock_choice')
                if stock_option.startswith('Última'):
                    snap=recent.iloc[0]
                    physical=stock_details(snap.id)
                    stock_by_code=physical.set_index('codigo')['bultos'].to_dict()
                    grouped=source_pending.groupby('CODIGO',as_index=False).agg(DESCRIPCION=('DESCRIPCION','first'),PENDIENTE=('PENDIENTE','sum'))
                    grouped['STOCK RESTANTE']=grouped.CODIGO.map(stock_by_code).fillna(0)
                    grouped['NECESIDAD NETA']=(grouped.PENDIENTE-grouped['STOCK RESTANTE']).clip(lower=0)
                    needs=grouped[grouped['NECESIDAD NETA']>0].copy()
                    st.caption(f'Se usa fotografía Bejerman #{int(snap.id)} del {snap.fecha_stock}; no se descuenta de nuevo la distribución.')
                latest_formula=versions.sort_values('version').drop_duplicates('codigo',keep='last')
                formula_by_code=latest_formula.set_index('codigo').to_dict('index')
                formula_options=['— Sin fórmula —']+latest_formula['codigo'].astype(str).tolist()
                selected_formulas={}
                valid_prices=prices[prices.ESTADO=='Válido'].copy()
                # Explicit mapping: profitability file has names but no product codes.
                st.subheader('1. Vincular productos con costos y fórmulas')
                st.caption('Vinculá manualmente cada código de Distribución con su fórmula y su precio. No se presupone que los códigos internos de fórmulas coincidan con los de Bejerman.')
                options=['— Sin vincular —']+valid_prices.PRODUCTO.tolist()
                units_per_bulto=st.number_input('Packs del rendimiento Bejerman por cada bulto comercial',min_value=1,value=1,step=1,key='v11_packs_per_bulto',help='Si Bejerman expresa packs y cada bulto comercial contiene varios packs, ingresá cuántos packs forman un bulto. Si son equivalentes, dejá 1.')
                mappings=[]
                for _, row in needs.iterrows():
                    code=str(row["CODIGO"])
                    with st.expander(f'{code} · {row["DESCRIPCION"]} · faltan {row["NECESIDAD NETA"]:,.0f} bultos',expanded=False):
                        formula_code=st.selectbox('Fórmula vigente',formula_options,index=formula_options.index(code) if code in formula_by_code else 0,format_func=lambda x: '— Sin fórmula —' if x=='— Sin fórmula —' else f'{formula_by_code[x]["producto"]} ({x})',key=f'v11_formula_{code}')
                        if formula_code!='— Sin fórmula —':selected_formulas[code]=formula_code
                        selected=st.selectbox('Producto en Excel de rentabilidad',options,key=f'v11_match_{code}')
                    if selected!='— Sin vincular —' and code in selected_formulas:
                        mappings.append((code,selected))
                duplicate_prices=pd.Series([x[1] for x in mappings]).duplicated().any() if mappings else False
                duplicate_formulas=pd.Series(list(selected_formulas.values())).duplicated().any() if selected_formulas else False
                if duplicate_formulas:st.error('Una misma fórmula fue vinculada a varios códigos. Revisá antes de optimizar.')
                if duplicate_prices:st.error('Hay productos de rentabilidad vinculados a más de un código. Corregí las relaciones antes de optimizar.')
                missing_formula=needs[~needs.CODIGO.isin(selected_formulas)]
                if not missing_formula.empty:
                    st.warning(f'{len(missing_formula)} productos con faltantes no tienen fórmula vigente y quedan fuera del plan.')
                    with st.expander('Ver productos sin fórmula'):
                        st.dataframe(missing_formula,hide_index=True,use_container_width=True)
                mapped=[]
                for code,product in mappings:
                    item=needs.loc[needs.CODIGO==code].iloc[0]
                    cost_row=valid_prices.loc[valid_prices.PRODUCTO==product].iloc[0]
                    formula=formula_by_code[selected_formulas[code]]
                    if pd.isna(formula['rendimiento_bejerman']) or float(formula['rendimiento_bejerman']) < units_per_bulto:
                        st.warning(f'{code}: rendimiento Bejerman pendiente o insuficiente. Se excluye del plan.')
                        continue
                    mapped.append({'codigo':code,'descripcion':str(item.DESCRIPCION),'demanda_neta':float(item['NECESIDAD NETA']),
                        'rendimiento':float(formula['rendimiento_bejerman'])/units_per_bulto,
                        'costo_bulto':float(cost_row['COSTO POR BULTO']),'precio_bulto':float(cost_row['PRECIO POR BULTO'])})
                st.subheader('2. Presupuesto para fabricar')
                financial_budget=0.
                if 'analysis' in st.session_state:
                    fc,_,cc,rr,hh,res,start=st.session_state['analysis']
                    end=start+pd.Timedelta(days=hh)
                    dated=fc.copy();dated['FECHA COBRANZA']=pd.to_datetime(dated['FECHA COBRANZA'])
                    usable=dated[dated['FECHA COBRANZA'].between(start,end)]
                    financial_budget=max(0.,float(cc)+float(usable.IMPORTE.sum())*(1-float(rr)/100)-float(res)-float(st.session_state.get('cc_obligation_reserve',0.)))
                    st.caption(f'Presupuesto proyectado hasta {end:%d/%m/%Y}, descontando riesgo, reserva de seguridad y obligaciones seleccionadas ({money(st.session_state.get("cc_obligation_reserve",0.))}). Las cobranzas futuras no son caja confirmada.')
                else:
                    st.warning('Todavía no analizaste cobranzas. Podés usar un presupuesto manual o ir a «Cobranzas y Cash Flow».')
                budget_mode=st.radio('Origen del presupuesto',['Proyección de cobranzas','Presupuesto manual'],horizontal=True,key='v11_budget_mode')
                manual=st.number_input('Presupuesto manual ($)',min_value=0.,value=0.,step=100000.,key='v11_manual_budget') if budget_mode=='Presupuesto manual' else 0.
                if budget_mode=='Proyección de cobranzas' and 'analysis' not in st.session_state:
                    st.warning('Analizá cobranzas o elegí presupuesto manual.')
                else:
                    gross_budget=financial_budget if budget_mode=='Proyección de cobranzas' else manual
                    expenses=st.number_input('Egresos comprometidos antes de fabricar ($)',min_value=0.,value=0.,step=100000.,key='v11_expenses',help='Descontá proveedores, salarios, impuestos y otras obligaciones que no estén incluidas en el presupuesto.')
                    net_budget=max(0.,gross_budget-expenses)
                    st.metric('Presupuesto máximo para producción',money(net_budget))
                    st.caption('Es un techo financiero, no una orden de compra. No contempla plazos de elaboración ni disponibilidad de insumos.')
                    if st.button('⚙️ Optimizar producción',type='primary',use_container_width=True,key='v11_optimize',disabled=not mapped or duplicate_prices or duplicate_formulas):
                        plan_result,issues,summary=optimize_batches(pd.DataFrame(mapped),net_budget)
                        st.session_state['v11_result']=(plan_result,issues,summary,net_budget,source_date,len(mapped),sum(float(x['demanda_neta']) for x in mapped))
                    if 'v11_result' in st.session_state:
                        plan_result,issues,summary,used_budget,used_date,mapped_count,mapped_demand=st.session_state['v11_result']
                        st.subheader('3. Recomendación más rentable')
                        st.caption(f'Resultado calculado con presupuesto {money(used_budget)} y distribución {used_date}. Presioná Optimizar nuevamente si modificás entradas.')
                        a,b,c,d=st.columns(4)
                        a.metric('Elaboraciones',str(summary.get('selected_batches',0)))
                        b.metric('Inversión',money(summary['investment']))
                        c.metric('Ganancia potencial',money(summary['profit']))
                        d.metric('Saldo del presupuesto',money(max(0.,used_budget-summary['investment'])))
                        st.caption(f'Estado: {summary["status"]}. La ganancia potencial descuenta el costo de toda la elaboración, incluso los bultos excedentes; solo reconoce ingresos por demanda respaldada.')
                        if not plan_result.empty:
                            st.dataframe(plan_result,hide_index=True,use_container_width=True)
                            st.metric('Bultos pendientes aún sin cubrir entre los productos vinculados',f'{max(0,mapped_demand-plan_result["BULTOS VENDIBLES"].sum()):,.0f}')
                            if plan_result['EXCEDENTE'].sum()>0:st.warning('El plan genera excedentes que quedan en stock y no se cuentan como ventas.')
                            buff=io.BytesIO()
                            with pd.ExcelWriter(buff,engine='xlsxwriter') as writer:
                                plan_result.to_excel(writer,sheet_name='Elaboraciones recomendadas',index=False)
                                pd.DataFrame(mapped).to_excel(writer,sheet_name='Datos de entrada',index=False)
                                if not issues.empty:issues.to_excel(writer,sheet_name='Datos a revisar',index=False)
                            st.download_button('📥 Descargar plan de producción',buff.getvalue(),file_name='plan_produccion_cc_v11.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                        else:
                            st.warning('Con los costos, demanda y presupuesto disponibles, no se recomienda ninguna elaboración rentable y financiable.')
                        if not issues.empty:st.dataframe(issues,hide_index=True,use_container_width=True)
        except Exception as exc:
            st.error('No se pudo preparar el plan: '+str(exc))
    st.warning('Plan preliminar: no se consideran capacidad, plazos, materias primas ni gastos no cargados. El historial SQLite local no es persistente en Streamlit Cloud.')
