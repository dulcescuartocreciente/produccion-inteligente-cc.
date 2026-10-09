import io
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Producción Inteligente CC', page_icon='🏭', layout='wide')
st.markdown('''<style>
.block-container{padding-top:1.3rem;max-width:1300px}
[data-testid="stMetric"]{background:#f1f8fd;border:1px solid #d8e9f6;border-radius:12px;padding:14px}
[data-testid="stMetricLabel"]{color:#24445c}
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


def read_route(file):
    book = pd.ExcelFile(file)
    if '9102026' not in book.sheet_names:
        raise ValueError('No se encontró la pestaña 9102026. Revisá que sea la hoja de ruta del 9 de octubre.')
    raw = pd.read_excel(book, sheet_name='9102026', header=None)
    records = []
    current_date = pd.Timestamp('2026-10-09')
    active = None
    for i in range(len(raw)):
        values = raw.iloc[i].tolist()
        row = [clean(v) for v in values]
        # Cada bloque puede tener su propio encabezado y distinta escritura.
        client_col = next((j for j, v in enumerate(row) if v == 'CLIENTE'), None)
        amount_col = next((j for j, v in enumerate(row) if v == 'IMPORTE'), None)
        pay_col = next((j for j, v in enumerate(row) if v.startswith('COBRAR')), None)
        if client_col is not None and amount_col is not None and pay_col is not None:
            bultos_col = next((j for j, v in enumerate(row) if v.startswith('BULTOS')), None)
            active = (client_col, amount_col, pay_col, bultos_col)
            continue
        # La fila con fecha y VUELTA marca el inicio de un bloque.
        if len(values) > 1 and 'VUELTA' in row[1]:
            active = None
            continue
        if active is None:
            continue
        c, a, pay, bultos = active
        client = values[c]
        if pd.isna(client) or not str(client).strip():
            continue
        client_name = str(client).strip()
        if any(word in clean(client_name) for word in ('TOTAL', 'CANCELADO', 'ANULADO')):
            continue
        # Ignorar subtítulos y notas; conservar filas con clientes y datos operativos.
        amount = parse_amount(values[a])
        condition = values[pay]
        days, state = payment_days(condition)
        bultos_val = values[bultos] if bultos is not None else np.nan
        if pd.isna(amount) and pd.isna(bultos_val) and pd.isna(condition):
            continue
        records.append({'FECHA': current_date, 'CLIENTE': client_name,
                        'IMPORTE': amount, 'BULTOS': bultos_val,
                        'CONDICIÓN DE PAGO': '' if pd.isna(condition) else str(condition).strip(),
                        'DÍAS DE PAGO': days, 'ESTADO CONDICIÓN': state})
    if not records:
        raise ValueError('No se encontraron entregas en la pestaña 9102026. Verificá los encabezados CLIENTE, COBRAR e IMPORTE.')
    data = pd.DataFrame(records)
    data['CLAVE CLIENTE'] = data['CLIENTE'].map(normalize_client)
    return data.reset_index(drop=True)


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
    return '$ ' + f'{v:,.0f}'.replace(',','.')

with st.sidebar:
    st.header('Archivos del cálculo')
    route_file = st.file_uploader('Hoja de ruta / importes (.xlsx)', type=['xlsx','xls'], key='ruta', help='Subí el Excel de la hoja de ruta. Se leerá únicamente la pestaña 9102026, con CLIENTE, IMPORTE y COBRAR (días).')
    history_file = st.file_uploader('Flujo histórico de cobranzas (.xlsx)', type=['xlsx','xls'], key='historico', help='Opcional. Subí el Excel con pedidos ya entregados y fechas de cobranza. Se muestra por separado y no se suma automáticamente a las ventas de la ruta para evitar duplicaciones.')
    st.divider()
    st.header('Parámetros')
    initial_cash = st.number_input('Caja disponible inicial ($)', min_value=0.0, value=0.0, step=100000.0, format='%.2f', help='Dinero efectivamente disponible hoy. No incluyas ventas que todavía no cobraste.')
    risk = st.slider('Margen de cobranza no realizada (%)', 0, 60, 20, help='Porcentaje que se descuenta de las cobranzas previstas por riesgo de demora o incumplimiento. No representa cobros reales.')
    horizon = st.selectbox('Horizonte de planificación', [7,15,30,60,90], index=2, format_func=lambda x:f'{x} días', help='Cantidad de días hacia adelante que se incluyen en la proyección, a partir de la fecha inicial elegida.')
    safety = st.number_input('Reserva de seguridad ($)', min_value=0.0, value=0.0, step=100000.0, format='%.2f', help='Dinero que querés reservar y no considerar dentro del techo preliminar para producir.')
    reference_date = st.date_input('Fecha inicial de planificación', value=datetime.now().date(), format='DD/MM/YYYY', help='Primer día desde el cual se cuentan las cobranzas dentro del horizonte. Para probar la ruta del 9 de octubre, elegí 09/10/2026.')
    st.caption('Los pagos y egresos comprometidos se incorporarán en la próxima etapa.')

if not route_file:
    st.subheader('Comenzá cargando tu hoja de ruta')
    st.write('Subí la hoja de ruta. Si contiene la pestaña **9102026**, se procesará exclusivamente el 9 de octubre. El flujo histórico es opcional.')
    st.stop()

try:
    route = read_route(route_file)
    history, payments = read_historical(history_file) if history_file else (pd.DataFrame(), pd.DataFrame())
except Exception as e:
    st.error(f'No se pudieron leer los archivos: {e}')
    st.stop()

if route.empty:
    st.warning('La hoja de ruta no contiene entregas con importe positivo.')
    st.stop()

st.subheader('1. Condiciones de pago de la hoja de ruta')
st.write('La condición se lee directamente de **COBRAR** o **COBRAR (DIAS)**. Los valores 0, CONTADO, RETIRAR PAGO y ANTICIPADO se interpretan como cobranza en la fecha de entrega. Los plazos numéricos se suman a esa fecha. No hace falta confirmar las condiciones reconocidas.')
terms = route[['FECHA', 'CLIENTE', 'CONDICIÓN DE PAGO', 'DÍAS DE PAGO', 'ESTADO CONDICIÓN', 'IMPORTE']].copy()
st.dataframe(terms, hide_index=True, use_container_width=True)
missing = route[route['ESTADO CONDICIÓN'] != 'AUTOMÁTICA']
if not missing.empty:
    st.warning(f'{len(missing)} fila(s) necesitan una fecha de cobranza o condición más precisa. No se incluirán en la proyección automática.')

st.subheader('2. Calendario de cobranzas proyectadas')
forecast = route.copy()
forecast['FECHA COBRANZA'] = forecast['FECHA'] + pd.to_timedelta(forecast['DÍAS DE PAGO'], unit='D')
forecast['COBRANZA AJUSTADA'] = forecast['IMPORTE'] * (1-risk/100)
forecast['ESTADO'] = np.where(forecast['ESTADO CONDICIÓN'].eq('AUTOMÁTICA') & forecast['IMPORTE'].notna() & forecast['IMPORTE'].gt(0), 'PROYECTABLE', 'REVISAR')
start = pd.Timestamp(reference_date)
end = start + pd.Timedelta(days=horizon)
usable = forecast[(forecast['ESTADO']=='PROYECTABLE') & forecast['FECHA COBRANZA'].between(start,end)].copy()
expected = float(usable['IMPORTE'].sum())
adjusted = float(usable['COBRANZA AJUSTADA'].sum())
budget = max(0.0, float(initial_cash)+adjusted-float(safety))
a,b,c,d = st.columns(4)
a.metric('Ventas de la hoja de ruta',money(route.IMPORTE.sum(skipna=True)))
b.metric(f'Cobranzas previstas ({horizon} días)',money(expected))
c.metric('Cobranzas ajustadas',money(adjusted))
d.metric('Techo preliminar de producción',money(budget))
st.caption('El techo es preliminar: todavía NO descuenta sueldos, proveedores, impuestos, otros pagos ni necesidades de capital de trabajo. No equivale a dinero libre para gastar. Los importes sin condición reconocida no se incluyen.')
if not usable.empty:
    daily = usable.groupby('FECHA COBRANZA',as_index=False).agg(PREVISTO=('IMPORTE','sum'),AJUSTADO=('COBRANZA AJUSTADA','sum'))
    st.bar_chart(daily.set_index('FECHA COBRANZA')[['PREVISTO','AJUSTADO']])
else:
    st.info('No hay cobranzas con condiciones confirmadas dentro del período seleccionado.')

t1,t2,t3,t4 = st.tabs(['Detalle de cobranzas','Resumen diario','Clientes y condiciones','Histórico'])
with t1:
    st.dataframe(forecast[['FECHA','CLIENTE','BULTOS','IMPORTE','CONDICIÓN DE PAGO','DÍAS DE PAGO','FECHA COBRANZA','COBRANZA AJUSTADA','ESTADO']] if 'BULTOS' in forecast else forecast[['FECHA','CLIENTE','IMPORTE','CONDICIÓN DE PAGO','DÍAS DE PAGO','FECHA COBRANZA','COBRANZA AJUSTADA','ESTADO']],hide_index=True,use_container_width=True)
with t2:
    if not usable.empty:
        st.dataframe(daily,hide_index=True,use_container_width=True)
with t3:
    st.dataframe(terms,hide_index=True,use_container_width=True)
with t4:
    if history.empty:st.info('No cargaste un archivo histórico.')
    else:
        st.metric('Operaciones históricas con importe',len(history))
        st.metric('Registros de cobros fechados detectados',len(payments))
        st.dataframe(history[['FECHA ENTREGA','CLIENTE','IMPORTE']].tail(100),hide_index=True,use_container_width=True)
        if payments.empty:st.warning('No se detectaron importes de cobranza en las columnas de fechas. No se infieren automáticamente condiciones de pago.')

output = io.BytesIO()
with pd.ExcelWriter(output,engine='xlsxwriter',datetime_format='dd/mm/yyyy') as writer:
    forecast.drop(columns=['CLAVE CLIENTE']).to_excel(writer,sheet_name='Proyección cobranzas',index=False)
    if not usable.empty:daily.to_excel(writer,sheet_name='Resumen diario',index=False)
    terms.to_excel(writer,sheet_name='Condiciones clientes',index=False)
    pd.DataFrame({'CONCEPTO':['Caja inicial','Cobranzas ajustadas','Reserva seguridad','Techo preliminar sin egresos'], 'IMPORTE':[initial_cash,adjusted,safety,budget]}).to_excel(writer,sheet_name='Resumen financiero',index=False)
st.download_button('📥 Descargar planificación financiera en Excel',output.getvalue(),file_name='produccion_cc_proyeccion_cobranzas.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary', help='Descargá un Excel con las fechas de cobro estimadas, el detalle por cliente, el resumen diario y los parámetros financieros.')
st.caption('Próximos módulos: egresos, costos y rentabilidad, recetas, materias primas y plan óptimo de fabricación.')
