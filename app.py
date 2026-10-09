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
    nums = pd.to_numeric(series, errors='coerce')
    as_num = pd.to_datetime(nums, unit='D', origin='1899-12-30', errors='coerce')
    as_text = pd.to_datetime(series.where(nums.isna()), errors='coerce', dayfirst=True)
    return as_num.fillna(as_text).dt.normalize()


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
    book = pd.ExcelFile(file)
    for sheet in book.sheet_names:
        raw = pd.read_excel(book, sheet_name=sheet, header=None)
        for i in range(min(20,len(raw))):
            row = [clean(v) for v in raw.iloc[i].tolist()]
            if 'CLIENTE' in row and 'IMPORTE' in row and ('FECHA ENTREGA' in row or 'FECHA' in row):
                cols = {key: row.index(key) for key in ('CLIENTE','IMPORTE')}
                cols['FECHA'] = row.index('FECHA ENTREGA') if 'FECHA ENTREGA' in row else row.index('FECHA')
                data = raw.iloc[i+1:].copy()
                out = pd.DataFrame({'FECHA ENTREGA': excel_date(data.iloc[:,cols['FECHA']]),
                                    'CLIENTE':data.iloc[:,cols['CLIENTE']].astype('string').str.strip(),
                                    'IMPORTE':pd.to_numeric(data.iloc[:,cols['IMPORTE']],errors='coerce')})
                out = out.dropna(subset=['FECHA ENTREGA','CLIENTE','IMPORTE'])
                out = out[(out.CLIENTE!='') & (out.IMPORTE>0)].copy()
                out['CLAVE CLIENTE'] = out.CLIENTE.map(normalize_client)
                # Fechas en columnas E en adelante; importes efectivamente cargados en celdas.
                dates = []
                for j in range(4, raw.shape[1]):
                    d = excel_date(pd.Series([raw.iat[i,j]])).iloc[0]
                    if pd.notna(d):
                        dates.append((j,d))
                records = []
                for idx,r in out.iterrows():
                    for j,d in dates:
                        amount = pd.to_numeric(pd.Series([raw.iat[idx,j]]),errors='coerce').iloc[0]
                        if pd.notna(amount) and amount>0:
                            records.append({'CLAVE CLIENTE':r['CLAVE CLIENTE'], 'FECHA ENTREGA':r['FECHA ENTREGA'], 'FECHA COBRANZA':d,'IMPORTE COBRANZA':float(amount)})
                return out.reset_index(drop=True), pd.DataFrame(records)
    raise ValueError('No encontré FECHA ENTREGA, CLIENTE e IMPORTE en el flujo histórico.')


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
    history_file=st.file_uploader('Cash Flow / flujo de cobranzas (.xlsx)',type=['xlsx','xls'],key='cash_flow',help='Subí el Excel de pedidos entregados y condiciones de pago. Se analiza por separado de las rutas para evitar sumar dos veces la misma venta.')
    st.divider()
    st.header('Parámetros')
    initial_cash=st.number_input('Caja disponible inicial ($)',min_value=0.0,value=0.0,step=100000.0,help='Dinero que ya está disponible. No incluye ventas pendientes de cobro.')
    risk=st.slider('Margen de cobranza no realizada (%)',0,60,20,help='Porcentaje de cobranzas previstas que se descuenta por precaución.')
    horizon=st.selectbox('Horizonte de planificación',[7,15,30,60,90],index=2,format_func=lambda x:f'{x} días',help='Días a considerar desde la fecha inicial.')
    safety=st.number_input('Reserva de seguridad ($)',min_value=0.0,value=0.0,step=100000.0,help='Dinero que preferís no comprometer en producción.')
    reference_date=st.date_input('Fecha inicial de planificación',value=datetime.now().date(),format='DD/MM/YYYY',help='Fecha desde la cual se analizan las cobranzas proyectadas.')
    st.caption('Los egresos se incorporarán más adelante.')

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

if st.button('🔎 Analizar cobranzas',type='primary',use_container_width=True,help='Calcula la proyección con las rutas guardadas y analiza el Cash Flow cargado por separado.'):
    historical_data=pd.DataFrame();historical_payments=pd.DataFrame();historical_error=None
    if history_file is not None:
        try:
            historical_data,historical_payments=read_historical(history_file)
        except Exception as exc:
            historical_error=str(exc)
    st.session_state['analysis']=(saved.copy(),float(initial_cash),int(risk),int(horizon),float(safety),pd.Timestamp(reference_date),historical_data,historical_payments,historical_error)

if 'analysis' in st.session_state:
    route,cash,risk_used,horizon_used,reserve,start,history,payments,history_error=st.session_state['analysis']
    if history_error:st.error('No se pudo leer el Cash Flow: '+history_error)
    with st.expander('Cash Flow histórico de cobranzas',expanded=not history.empty):
        st.caption('Se presenta por separado para no duplicar importes que también puedan figurar en las hojas de ruta. No se suma automáticamente al techo preliminar.')
        if history.empty:
            st.info('Para ver el Cash Flow, cargá su Excel en el panel izquierdo y presioná Analizar cobranzas.')
        else:
            a_hist,b_hist=st.columns(2)
            a_hist.metric('Entregas en Cash Flow',len(history))
            b_hist.metric('Cobranzas fechadas detectadas',len(payments))
            st.dataframe(history,hide_index=True,use_container_width=True)
            if not payments.empty:
                st.dataframe(payments,hide_index=True,use_container_width=True)
    if route.empty:
        st.info('Todavía no hay hojas de ruta guardadas. Podés consultar el Cash Flow arriba; para proyectar rutas, guardá una hoja y volvé a analizar.')
    else:
        forecast=route.copy()
        forecast['FECHA COBRANZA']=forecast['FECHA']+pd.to_timedelta(forecast['DÍAS DE PAGO'],unit='D')
        forecast['COBRANZA AJUSTADA']=forecast['IMPORTE']*(1-risk_used/100)
        forecast['ESTADO']=np.where(forecast['ESTADO CONDICIÓN'].eq('AUTOMÁTICA') & forecast['IMPORTE'].notna() & forecast['IMPORTE'].gt(0),'PROYECTABLE','REVISAR')
        usable=forecast[(forecast['ESTADO']=='PROYECTABLE') & forecast['FECHA COBRANZA'].between(start,start+pd.Timedelta(days=horizon_used))]
        expected=float(usable['IMPORTE'].sum());adjusted=float(usable['COBRANZA AJUSTADA'].sum())
        budget=max(0,cash+adjusted-reserve)
        with indicators.container():
            a,b,c,d=st.columns(4)
            a.metric('Ventas registradas',money(float(route['IMPORTE'].sum())))
            b.metric(f'Cobranzas ({horizon_used} días)',money(expected))
            c.metric('Cobranzas ajustadas',money(adjusted))
            d.metric('Techo preliminar',money(budget))
        st.caption('El techo preliminar no descuenta proveedores, sueldos ni impuestos; no representa dinero libre para gastar.')
        daily=usable.groupby('FECHA COBRANZA',as_index=False).agg(PREVISTO=('IMPORTE','sum'),AJUSTADO=('COBRANZA AJUSTADA','sum'))
        if not daily.empty:st.bar_chart(daily.set_index('FECHA COBRANZA'))
        else:st.info('No hay cobranzas proyectadas en el período seleccionado.')
        t1,t2,t3=st.tabs(['Detalle de cobranzas','Resumen diario','Revisar condiciones'])
        with t1:st.dataframe(forecast,hide_index=True,use_container_width=True)
        with t2:st.dataframe(daily,hide_index=True,use_container_width=True)
        with t3:st.dataframe(forecast[forecast['ESTADO']=='REVISAR'],hide_index=True,use_container_width=True)
        output=io.BytesIO()
        with pd.ExcelWriter(output,engine='xlsxwriter',datetime_format='dd/mm/yyyy') as writer:
            forecast.to_excel(writer,sheet_name='Proyeccion cobranzas',index=False)
            daily.to_excel(writer,sheet_name='Resumen diario',index=False)
            pd.DataFrame({'CONCEPTO':['Caja inicial','Cobranzas ajustadas','Reserva','Techo preliminar'],'IMPORTE':[cash,adjusted,reserve,budget]}).to_excel(writer,sheet_name='Resumen financiero',index=False)
        st.download_button('📥 Descargar análisis en Excel',output.getvalue(),file_name='produccion_cc_analisis.xlsx',help='Exporta el análisis y el calendario de cobranzas calculado.')
else:
    st.info('Cargá y guardá una hoja de ruta; después presioná **Analizar cobranzas**.')
