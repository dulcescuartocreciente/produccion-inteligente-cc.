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
st.info('Proyectá cobranzas, analizá rentabilidad y administrá fórmulas de fabricación. El plan automático quedará disponible al incorporar pedidos y stock.')


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
    profitability_file=st.file_uploader('Costos, márgenes y precios por bulto (.xlsx)',type=['xlsx','xls'],key='profitability',help='Subí el Excel con las columnas Artic.-Conc., COSTO2, MARGEN y precio. Los costos cero quedan excluidos de la simulación.')
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


st.divider()
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
                available = max(0.0, cc + float(within['IMPORTE'].sum()) * (1-rr/100) - res)
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


st.divider()
st.header('🧪 Fórmulas y rendimientos')
st.caption('El plan utiliza el rendimiento Bejerman (incluye merma). Solo se admiten elaboraciones completas. Cada edición genera una versión nueva y conserva las anteriores.')

def formulas_db():
    with sqlite3.connect(store_path()) as con:
        con.execute('CREATE TABLE IF NOT EXISTS formulas (codigo TEXT, version INTEGER, producto TEXT, rendimiento_bejerman REAL, rendimiento_teorico REAL, observaciones TEXT, creado TEXT, origen TEXT, PRIMARY KEY(codigo,version))')
        con.execute('CREATE TABLE IF NOT EXISTS formula_ingredientes (codigo TEXT, version INTEGER, posicion INTEGER, grupo TEXT, ingrediente TEXT, cantidad REAL, unidad TEXT, brix REAL, PRIMARY KEY(codigo,version,posicion))')

def formula_versions():
    with sqlite3.connect(store_path()) as con:
        return pd.read_sql_query('SELECT * FROM formulas ORDER BY producto,version DESC',con)

def formula_ingredients(code,version):
    with sqlite3.connect(store_path()) as con:
        return pd.read_sql_query('SELECT grupo,ingrediente,cantidad,unidad,brix FROM formula_ingredientes WHERE codigo=? AND version=? ORDER BY posicion',con,params=(code,int(version)))

def save_formula(code,product,bejerman,theoretical,notes,ingredients,origin):
    code=str(code).strip().upper();product=str(product).strip()
    if not code or not product: raise ValueError('Completá el código y el nombre del producto.')
    if not np.isfinite(bejerman) or bejerman<=0: raise ValueError('El rendimiento Bejerman debe ser mayor que cero.')
    if ingredients.empty: raise ValueError('La fórmula debe tener ingredientes.')
    data=[]
    for _,r in ingredients.iterrows():
        name=str(r.get('ingrediente','')).strip()
        if not name or name.lower()=='nan': continue
        qty=pd.to_numeric(r.get('cantidad'),errors='coerce')
        if pd.isna(qty) or qty<0: raise ValueError('Revisá las cantidades: deben ser números no negativos.')
        brix=pd.to_numeric(r.get('brix'),errors='coerce')
        data.append((str(r.get('grupo','') or ''),name,float(qty),str(r.get('unidad','') or ''),None if pd.isna(brix) else float(brix)))
    if not data: raise ValueError('No hay ingredientes válidos.')
    with sqlite3.connect(store_path()) as con:
        last=con.execute('SELECT MAX(version) FROM formulas WHERE codigo=?',(code,)).fetchone()[0]
        version=1 if last is None else last+1
        con.execute('INSERT INTO formulas VALUES (?,?,?,?,?,?,?,?)',(code,version,product,float(bejerman),None if pd.isna(theoretical) else float(theoretical),notes,datetime.now().isoformat(timespec='seconds'),origin))
        con.executemany('INSERT INTO formula_ingredientes VALUES (?,?,?,?,?,?,?,?)',[(code,version,i,*r) for i,r in enumerate(data)])
    return version

def parse_formula_excel(upload):
    book=pd.ExcelFile(upload)
    raw=pd.read_excel(book,sheet_name=0,header=None)
    def cell(r,c):
        return raw.iat[r,c] if r<len(raw) and c<len(raw.columns) else None
    code=None;name=None;revision=None;bej=None;theo=None;items=[];in_table=False;group=''
    for i in range(len(raw)):
        vals=[clean(v) if pd.notna(v) else '' for v in raw.iloc[i].tolist()]
        if 'VERSION' in vals and 'CODIGO' in vals:
            revision=cell(i+1,vals.index('VERSION'));code=cell(i+1,vals.index('CODIGO'))
        if any('MERMELADA' in v or 'JUGO' in v or 'DULCE' in v for v in vals[:2]) and name is None:
            name=cell(i,0)
        label=vals[0] if vals else ''
        if 'RENDIMIENTO BEJERMAN' in label:bej=pd.to_numeric(cell(i,2),errors='coerce')
        if label=='RENDIMIENTO TEORICO':theo=pd.to_numeric(cell(i,2),errors='coerce')
        if len(vals)>3 and vals[1]=='INGREDIENTE' and vals[2]=='CANTIDAD':in_table=True;continue
        if in_table:
            if 'MASA TOTAL' in label or 'DATOS FISICOQUIMICOS' in label:in_table=False;continue
            ingredient=cell(i,1);qty=pd.to_numeric(cell(i,2),errors='coerce')
            if pd.notna(cell(i,0)) and str(cell(i,0)).strip():group=str(cell(i,0)).strip()
            if pd.notna(ingredient) and str(ingredient).strip() and pd.notna(qty):
                brix=pd.to_numeric(cell(i,4),errors='coerce')
                items.append({'grupo':group,'ingrediente':str(ingredient).strip(),'cantidad':float(qty),'unidad':str(cell(i,3) or ''),'brix':None if pd.isna(brix) else float(brix)})
    if not code or pd.isna(bej) or not items:
        raise ValueError('No pude encontrar código, rendimiento Bejerman e ingredientes. Revisá que el archivo tenga el formato del modelo.')
    return {'codigo':str(code).strip(),'producto':str(name or code).strip(),'revision_original':revision,'bejerman':float(bej),'teorico':float(theo) if pd.notna(theo) else 0.,'ingredientes':pd.DataFrame(items)}

formulas_db()
f1,f2,f3=st.tabs(['📥 Importar fórmula','📚 Fórmulas e historial','✏️ Editar y simular'])
with f1:
    formula_file=st.file_uploader('Subir fórmula de elaboración (.xlsx)',type=['xlsx','xls'],key='formula_upload',help='Subí un Excel de fórmula como el modelo de Mermelada Stevia Arándano. Se leerán ingredientes, código y rendimiento Bejerman.')
    if formula_file:
        try:
            parsed=parse_formula_excel(formula_file)
            st.write(f"**{parsed['producto']}** · Código {parsed['codigo']} · Revisión del Excel: {parsed['revision_original']}")
            x,y=st.columns(2)
            x.metric('Rendimiento Bejerman (packs)',f"{parsed['bejerman']:,.2f}")
            y.metric('Rendimiento teórico (packs)',f"{parsed['teorico']:,.2f}")
            st.dataframe(parsed['ingredientes'],hide_index=True,use_container_width=True)
            import_note=st.text_input('Observaciones de importación',key='formula_import_note',help='Anotá el motivo de la carga o la revisión de la fórmula.')
            if st.button('Guardar fórmula como nueva versión',type='primary',key='formula_import_save'):
                version=save_formula(parsed['codigo'],parsed['producto'],parsed['bejerman'],parsed['teorico'],import_note,parsed['ingredientes'],f'Excel: {formula_file.name}; revisión origen {parsed["revision_original"]}')
                st.success(f'Fórmula guardada. Versión interna {version}. Se conservaron las versiones anteriores.')
        except Exception as exc:st.error('No se pudo importar la fórmula: '+str(exc))
with f2:
    versions=formula_versions()
    if versions.empty:st.info('Todavía no hay fórmulas guardadas. Importá la primera desde Excel.')
    else:
        latest=versions.sort_values('version').drop_duplicates('codigo',keep='last')
        st.subheader('Fórmulas vigentes')
        st.dataframe(latest[['codigo','producto','version','rendimiento_bejerman','rendimiento_teorico','creado']],hide_index=True,use_container_width=True)
        st.subheader('Historial de revisiones')
        st.dataframe(versions[['codigo','producto','version','rendimiento_bejerman','observaciones','creado','origen']],hide_index=True,use_container_width=True)
        export=io.BytesIO()
        with pd.ExcelWriter(export,engine='xlsxwriter') as writer:
            versions.to_excel(writer,sheet_name='Versiones',index=False)
            with sqlite3.connect(store_path()) as con:
                pd.read_sql_query('SELECT * FROM formula_ingredientes',con).to_excel(writer,sheet_name='Ingredientes',index=False)
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
            new_bej=st.number_input('Rendimiento Bejerman (packs por elaboración)',min_value=0.01,value=float(base['rendimiento_bejerman']),format='%.4f',help='Rendimiento con merma. Es el utilizado para calcular elaboraciones completas.')
            new_theo=st.number_input('Rendimiento teórico (referencia)',min_value=0.0,value=float(base['rendimiento_teorico'] or 0),format='%.4f')
            edited=st.data_editor(ingredients,num_rows='dynamic',hide_index=True,use_container_width=True,key=f'edit_{selected_code}_{chosen_version}',column_config={'cantidad':st.column_config.NumberColumn('Cantidad',min_value=0),'ingrediente':st.column_config.TextColumn('Ingrediente',required=True)})
            edit_note=st.text_input('Motivo del cambio / observaciones',help='El historial conserva esta nota junto con la fecha de la nueva versión.')
            submitted=st.form_submit_button('Guardar cambios como nueva versión',type='primary')
        if submitted:
            try:
                if not edit_note.strip():raise ValueError('Indicá el motivo de la modificación para conservar la trazabilidad.')
                new_ver=save_formula(selected_code,new_product,new_bej,new_theo,edit_note,edited,f'Edición de versión {chosen_version}')
                st.success(f'Nueva versión {new_ver} guardada. La versión {chosen_version} permanece en el historial.')
            except Exception as exc:st.error(str(exc))
        st.subheader('Simular necesidad de elaboraciones completas')
        demand=st.number_input('Bultos necesarios (después de descontar stock)',min_value=0,max_value=1000000,value=150,step=1,key='formula_demand')
        # Se descartan fracciones: no se pueden despachar packs incompletos.
        usable=int(np.floor(float(base['rendimiento_bejerman'])))
        if usable<1:st.warning('El rendimiento Bejerman es menor a un bulto completo. Revisá la unidad de medida.')
        else:
            batches=int(np.ceil(demand/usable))
            produced=batches*usable
            a,b,c=st.columns(3)
            a.metric('Elaboraciones completas',str(batches))
            b.metric('Bultos utilizables',str(produced))
            c.metric('Excedente estimado',str(produced-demand))
            st.caption(f'Se toman {usable} bultos completos por elaboración (rendimiento Bejerman {float(base["rendimiento_bejerman"]):.2f}). No se permiten elaboraciones parciales. Esta simulación no descuenta materias primas ni presupuesto.')

st.warning('Importante: el historial de fórmulas usa una base SQLite local. En Streamlit Cloud puede perderse al reiniciar o redesplegar. Descargá respaldos y conectemos una base de datos persistente antes de usarlo como registro definitivo.')
