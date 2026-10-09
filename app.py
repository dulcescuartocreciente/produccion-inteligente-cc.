import io
import hashlib
import sqlite3
import re
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Producción Inteligente CC', page_icon='🏭', layout='wide')
st.markdown('''<style>
.block-container{padding-top:1.3rem;max-width:1300px}
[data-testid="stMetric"]{background:#f1f8fd;border:1px solid #d8e9f6;border-radius:12px;padding:9px 10px;min-width:0}
[data-testid="stMetricLabel"]{color:#24445c}
[data-testid="stMetricValue"]{font-size:clamp(1rem,1.5vw,1.4rem)!important;overflow-wrap:anywhere;line-height:1.25}
[data-testid="stMetricLabel"]{font-size:.78rem!important}
h1,h2,h3{color:#175d8b}
</style>''', unsafe_allow_html=True)
logo = Path(__file__).with_name('logo_cc.png')
if logo.exists():
    st.image(str(logo), width=240)
st.title('🏭 Producción Inteligente CC')
st.caption('Planificación financiera de producción · Cuarto Creciente')
st.info('Esta primera etapa proyecta cobranzas y recursos estimados. Todavía no recomienda productos a fabricar: para eso incorporaremos costos, demanda, insumos y capacidad.')


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
    if term in ('CONTADO', 'EFECTIVO', 'PAGO CONTADO', 'ANTICIPADO', 'RETIRAR PAGO', 'RETIRA PAGO'):
        return 0, 'AUTOMÁTICA'
    return np.nan, 'CONDICIÓN NO RECONOCIDA'




def read_historical(file):
    """Importa cobros fechados del cash flow; nunca supone que IMPORTE ya se cobró."""
    book = pd.ExcelFile(file)
    records = []
    for sheet in book.sheet_names:
        raw = pd.read_excel(book, sheet_name=sheet, header=None)
        for i in range(min(25, len(raw))):
            labels = [clean(v) for v in raw.iloc[i].tolist()]
            if not ('CLIENTE' in labels and 'IMPORTE' in labels and
                    ('FECHA ENTREGA' in labels or 'FECHA' in labels)):
                continue
            c = labels.index('CLIENTE'); a = labels.index('IMPORTE')
            f = labels.index('FECHA ENTREGA') if 'FECHA ENTREGA' in labels else labels.index('FECHA')
            date_columns = []
            for j, v in enumerate(raw.iloc[i].tolist()):
                if j in (c,a,f): continue
                dt = pd.to_datetime(v, errors='coerce', dayfirst=True)
                if pd.notna(dt) and pd.Timestamp('2020-01-01') <= dt <= pd.Timestamp('2045-12-31'):
                    date_columns.append((j, pd.Timestamp(dt).normalize()))
            for k in range(i+1, len(raw)):
                row = raw.iloc[k]
                client = row.iloc[c]
                if pd.isna(client) or not str(client).strip(): continue
                delivery = excel_date(pd.Series([row.iloc[f]])).iloc[0]
                amount = parse_amount(row.iloc[a])
                if pd.isna(delivery) or pd.isna(amount) or amount <= 0: continue
                for j, date in date_columns:
                    payment = parse_amount(row.iloc[j])
                    if pd.notna(payment) and payment > 0:
                        records.append({'FECHA ENTREGA': delivery, 'CLIENTE': str(client).strip(),
                                        'IMPORTE VENTA': float(amount), 'FECHA COBRANZA': date,
                                        'IMPORTE': float(payment), 'ORIGEN': 'CASH FLOW'})
            break
    if not records:
        return pd.DataFrame(columns=['FECHA ENTREGA','CLIENTE','IMPORTE VENTA','FECHA COBRANZA','IMPORTE','ORIGEN'])
    return pd.DataFrame(records)


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


def money(v):
    return "$ " + f"{v:,.0f}".replace(",", ".")


def date_from_sheet(name):
    digits = re.sub(r'\D', '', str(name))
    for year_len in (4, 2):
        if not digits.endswith(('20' if year_len == 4 else '')):
            pass
        if len(digits) < year_len+2:
            continue
        year = int(digits[-year_len:])
        if year_len == 2: year += 2000
        front = digits[:-year_len]
        for day_len in (1,2):
            if len(front) <= day_len: continue
            try:
                day=int(front[:day_len]); month=int(front[day_len:])
                if not 1<=month<=12: continue
                return pd.Timestamp(year,month,day)
            except (ValueError, TypeError): pass
    return pd.NaT


def read_route(file):
    book = pd.ExcelFile(file)
    candidates=[]
    for name in book.sheet_names:
        raw=pd.read_excel(book,sheet_name=name,header=None)
        header_count=0
        for i in range(min(len(raw),40)):
            vals=[clean(v) for v in raw.iloc[i].tolist()]
            if 'CLIENTE' in vals and 'IMPORTE' in vals and any(v.startswith('COBRAR') for v in vals):
                header_count+=1
        if header_count: candidates.append((name,raw,header_count))
    if not candidates: raise ValueError('No encontré una hoja con CLIENTE, IMPORTE y COBRAR (días).')
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
        if c is not None and a is not None and pay is not None:
            b=next((j for j,v in enumerate(row) if v.startswith('BULTOS')),None)
            active=(c,a,pay,b);continue
        if len(row)>1 and 'VUELTA' in row[1]: active=None;continue
        if active is None: continue
        c,a,pay,b=active
        client=values[c]
        if pd.isna(client) or not str(client).strip():continue
        name=str(client).strip()
        if any(x in clean(name) for x in ('TOTAL','CANCELADO','ANULADO')):continue
        amount=parse_amount(values[a]); condition=values[pay]
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
    route_file=st.file_uploader('Hoja de ruta del día',type=['xlsx','xls'],help='Subí la hoja de ruta correspondiente al día que querés incorporar. Debe tener CLIENTE, IMPORTE y COBRAR (días). Si el libro contiene varios días, podrás elegir la pestaña.')
    history_file=st.file_uploader('Cash Flow / flujo de cobranzas (.xlsx)',type=['xlsx','xls'],key='cash_flow',help='Subí el Excel con las fechas e importes de cobranzas. Se sumará a la proyección evitando operaciones repetidas con las hojas de ruta.')
    st.divider()
    st.header('Parámetros')
    initial_cash=st.number_input('Caja disponible inicial ($)',min_value=0.0,value=0.0,step=100000.0,help='Dinero que ya está disponible. No incluye ventas pendientes de cobro.')
    risk=st.slider('Margen de cobranza no realizada (%)',0,60,20,help='Porcentaje de cobranzas previstas que se descuenta por precaución.')
    horizon=st.selectbox('Horizonte de planificación',[7,15,30,60,90],index=2,format_func=lambda x:f'{x} días',help='Días a considerar desde la fecha inicial.')
    safety=st.number_input('Reserva de seguridad ($)',min_value=0.0,value=0.0,step=100000.0,help='Dinero que preferís no comprometer en producción.')
    reference_date=st.date_input('Fecha inicial de planificación',value=datetime.now().date(),format='DD/MM/YYYY',help='Fecha desde la cual se analizan las cobranzas proyectadas.')
    st.caption('Los egresos se incorporarán más adelante. Las cobranzas proyectadas no equivalen a dinero efectivamente recibido.')

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
        forecast, excluded = build_forecast(saved, cashflow)
        st.session_state['analysis'] = (forecast, excluded, float(initial_cash), int(risk),
                                        int(horizon), float(safety), pd.Timestamp(reference_date))

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
        budget = max(0, cash + adjusted - reserve)
        with indicators.container():
            a,b,c,d = st.columns(4)
            a.metric('Caja inicial', money(cash))
            b.metric(f'Cobranzas ({horizon_used} días)', money(expected))
            c.metric('Ingresos prudentes', money(adjusted))
            d.metric('Techo preliminar', money(budget))
        st.caption('La caja proyectada es una simulación: las cobranzas futuras no están confirmadas. No incluye pagos a proveedores, sueldos ni impuestos.')
        if excluded:
            st.info(f'Se excluyeron {excluded} operaciones de hojas de ruta porque ya estaban en el Cash Flow (mismo cliente, fecha de entrega e importe).')
        st.subheader('Calendario de ingresos para planificar producción')
        daily = usable.groupby('FECHA COBRANZA',as_index=False).agg(
            COBRANZA_PREVISTA=('IMPORTE','sum'), COBRANZA_AJUSTADA=('COBRANZA AJUSTADA','sum'))
        days = pd.DataFrame({'FECHA COBRANZA':pd.date_range(start,end,freq='D')})
        daily = days.merge(daily,on='FECHA COBRANZA',how='left').fillna(0)
        daily['CAJA PROYECTADA'] = cash + daily['COBRANZA_AJUSTADA'].cumsum()
        daily['TECHO PRELIMINAR'] = (daily['CAJA PROYECTADA']-reserve).clip(lower=0)
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
            pd.DataFrame({'CONCEPTO':['Caja inicial','Ingresos previstos','Ingresos prudentes','Reserva','Techo preliminar'],
                          'IMPORTE':[cash,expected,adjusted,reserve,budget]}).to_excel(writer,sheet_name='Resumen financiero',index=False)
        st.download_button('📥 Descargar proyección de ingresos',output.getvalue(),
                           file_name='produccion_cc_proyeccion_ingresos.xlsx',
                           help='Descarga las fechas, importes y caja acumulada para planificar producción.')
else:
    st.info('Cargá el Cash Flow y/o guardá una hoja de ruta. Después presioná **Analizar cobranzas**.')
