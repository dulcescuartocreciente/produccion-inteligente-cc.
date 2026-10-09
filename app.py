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


def read_route(file):
    book = pd.ExcelFile(file)
    for sheet in book.sheet_names:
        raw = pd.read_excel(book, sheet_name=sheet, header=None)
        for i in range(min(20, len(raw))):
            header = [clean(x) for x in raw.iloc[i].tolist()]
            if {'FECHA','CLIENTE','IMPORTE'}.issubset(set(header)):
                raw.columns = header
                data = raw.iloc[i+1:][['FECHA','CLIENTE','IMPORTE'] + (['BULTOS'] if 'BULTOS' in header else [])].copy()
                data['FECHA'] = excel_date(data['FECHA'])
                data['CLIENTE'] = data['CLIENTE'].astype('string').str.strip()
                data['IMPORTE'] = pd.to_numeric(data['IMPORTE'], errors='coerce')
                data = data.dropna(subset=['FECHA','CLIENTE','IMPORTE'])
                data = data[(data['CLIENTE']!='') & (data['IMPORTE']>0)].copy()
                data['CLAVE CLIENTE'] = data['CLIENTE'].map(normalize_client)
                return data.reset_index(drop=True)
    raise ValueError('No encontré FECHA, CLIENTE e IMPORTE en la hoja de ruta.')


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
    route_file = st.file_uploader('Hoja de ruta / importes (.xlsx)', type=['xlsx','xls'], key='ruta')
    history_file = st.file_uploader('Flujo histórico de cobranzas (.xlsx)', type=['xlsx','xls'], key='historico')
    st.divider()
    st.header('Parámetros')
    initial_cash = st.number_input('Caja disponible inicial ($)', min_value=0.0, value=0.0, step=100000.0, format='%.2f')
    risk = st.slider('Margen de cobranza no realizada (%)', 0, 60, 20)
    horizon = st.selectbox('Horizonte de planificación', [7,15,30,60,90], index=2, format_func=lambda x:f'{x} días')
    safety = st.number_input('Reserva de seguridad ($)', min_value=0.0, value=0.0, step=100000.0, format='%.2f')
    reference_date = st.date_input('Fecha inicial de planificación', value=datetime.now().date(), format='DD/MM/YYYY')
    st.caption('Los pagos y egresos comprometidos se incorporarán en la próxima etapa.')

if not route_file:
    st.subheader('Comenzá cargando tu hoja de ruta')
    st.write('Usá el panel izquierdo para subir el Excel **IMPORTES**. El flujo histórico es opcional, pero permite consultar los clientes anteriores.')
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

st.subheader('1. Condiciones de pago por cliente')
st.write('Ingresá los días de pago de cada cliente. **El 100% del importe se asigna a una sola fecha**, contando desde la entrega. Podés modificar los días antes de calcular.')
clients = route[['CLAVE CLIENTE','CLIENTE']].drop_duplicates('CLAVE CLIENTE').sort_values('CLIENTE').reset_index(drop=True)
previous = set(history['CLAVE CLIENTE']) if not history.empty else set()
clients['EN HISTÓRICO'] = clients['CLAVE CLIENTE'].isin(previous)
clients['DÍAS DE PAGO'] = 0
clients['CONDICIÓN CONFIRMADA'] = False
st.caption('Por seguridad, los días comienzan en 0 y **ninguna condición está confirmada**. El histórico no permite deducir de forma confiable las condiciones si no hay importes de cobranza identificables. Confirmá cada cliente antes de proyectar.')
terms = st.data_editor(clients[['CLIENTE','DÍAS DE PAGO','CONDICIÓN CONFIRMADA','EN HISTÓRICO']], hide_index=True, use_container_width=True,
    disabled=['CLIENTE','EN HISTÓRICO'], num_rows='fixed', key='terms',
    column_config={'DÍAS DE PAGO':st.column_config.NumberColumn(min_value=0,max_value=365,step=1),
                   'CONDICIÓN CONFIRMADA':st.column_config.CheckboxColumn(help='Marcá solo cuando hayas verificado los días de pago')})
terms['CLAVE CLIENTE'] = clients['CLAVE CLIENTE'].values
valid = terms['CONDICIÓN CONFIRMADA'].fillna(False) & pd.to_numeric(terms['DÍAS DE PAGO'],errors='coerce').ge(0)
missing = terms.loc[~valid,'CLIENTE'].tolist()
if missing:
    st.warning(f'Faltan confirmar condiciones de pago para {len(missing)} cliente(s). Sus importes quedarán fuera de la proyección. Ejemplos: {", ".join(map(str,missing[:5]))}')

st.subheader('2. Calendario de cobranzas proyectadas')
forecast = route.merge(terms[['CLAVE CLIENTE','DÍAS DE PAGO','CONDICIÓN CONFIRMADA']],on='CLAVE CLIENTE',how='left')
forecast['FECHA COBRANZA'] = forecast['FECHA'] + pd.to_timedelta(pd.to_numeric(forecast['DÍAS DE PAGO'],errors='coerce').fillna(0),unit='D')
forecast['COBRANZA AJUSTADA'] = forecast['IMPORTE'] * (1-risk/100)
forecast['ESTADO'] = np.where(forecast['CONDICIÓN CONFIRMADA'].fillna(False),'CONFIRMADA','SIN CONDICIÓN')
start = pd.Timestamp(reference_date)
end = start + pd.Timedelta(days=horizon)
usable = forecast[(forecast['ESTADO']=='CONFIRMADA') & forecast['FECHA COBRANZA'].between(start,end)].copy()
expected = float(usable['IMPORTE'].sum())
adjusted = float(usable['COBRANZA AJUSTADA'].sum())
budget = max(0.0, float(initial_cash)+adjusted-float(safety))
a,b,c,d = st.columns(4)
a.metric('Ventas de la hoja de ruta',money(route.IMPORTE.sum()))
b.metric(f'Cobranzas previstas ({horizon} días)',money(expected))
c.metric('Cobranzas ajustadas',money(adjusted))
d.metric('Techo preliminar de producción',money(budget))
st.caption('El techo es preliminar: todavía NO descuenta sueldos, proveedores, impuestos, otros pagos ni necesidades de capital de trabajo. No equivale a dinero libre para gastar.')
if not usable.empty:
    daily = usable.groupby('FECHA COBRANZA',as_index=False).agg(PREVISTO=('IMPORTE','sum'),AJUSTADO=('COBRANZA AJUSTADA','sum'))
    st.bar_chart(daily.set_index('FECHA COBRANZA')[['PREVISTO','AJUSTADO']])
else:
    st.info('No hay cobranzas con condiciones confirmadas dentro del período seleccionado.')

t1,t2,t3,t4 = st.tabs(['Detalle de cobranzas','Resumen diario','Clientes y condiciones','Histórico'])
with t1:
    st.dataframe(forecast[['FECHA','CLIENTE','BULTOS','IMPORTE','DÍAS DE PAGO','FECHA COBRANZA','COBRANZA AJUSTADA','ESTADO']] if 'BULTOS' in forecast else forecast[['FECHA','CLIENTE','IMPORTE','DÍAS DE PAGO','FECHA COBRANZA','COBRANZA AJUSTADA','ESTADO']],hide_index=True,use_container_width=True)
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
    terms.drop(columns=['CLAVE CLIENTE']).to_excel(writer,sheet_name='Condiciones clientes',index=False)
    pd.DataFrame({'CONCEPTO':['Caja inicial','Cobranzas ajustadas','Reserva seguridad','Techo preliminar sin egresos'], 'IMPORTE':[initial_cash,adjusted,safety,budget]}).to_excel(writer,sheet_name='Resumen financiero',index=False)
st.download_button('📥 Descargar planificación financiera en Excel',output.getvalue(),file_name='produccion_cc_proyeccion_cobranzas.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',type='primary')
st.caption('Próximos módulos: egresos, costos y rentabilidad, recetas, materias primas y plan óptimo de fabricación.')
