"""Optimización de elaboraciones completas, con ganancia sobre demanda respaldada."""
import math
import numpy as np
import pandas as pd
from scipy.optimize import milp, Bounds, LinearConstraint


def optimize_batches(products, budget):
    """products: codigo, demanda_neta, rendimiento, costo_bulto, precio_bulto.
    Maximiza ingreso por bultos vendidos menos costo de todos los bultos elaborados.
    No atribuye ventas a excedentes. Retorna plan y diagnóstico.
    """
    if budget < 0 or not math.isfinite(budget):
        raise ValueError('El presupuesto debe ser un importe válido no negativo.')
    candidates=[]; issues=[]; variables=[]
    for row in products.to_dict('records'):
        code=str(row['codigo']); need=float(row['demanda_neta']); yield_raw=float(row['rendimiento'])
        cost=float(row['costo_bulto']); price=float(row['precio_bulto'])
        if not all(math.isfinite(x) for x in (need,yield_raw,cost,price)) or need < 0 or yield_raw < 1 or cost <= 0 or price <= 0:
            issues.append({'Código':code,'Motivo':'Demanda, rendimiento, costo o precio inválidos'})
            continue
        y=math.floor(yield_raw)
        if need <= 0:continue
        count=math.ceil(need/y)
        # One binary decision per whole batch. Bound model size for reliable interactive use.
        if count>200:
            issues.append({'Código':code,'Motivo':'Más de 200 elaboraciones requeridas: revisar unidades o dividir horizonte'})
            continue
        candidates.append(row)
        for j in range(count):
            sold=max(0,min(y,need-j*y))
            variables.append({'codigo':code,'batch':j,'yield':y,'sold':sold,
                              'investment':y*cost,'profit':sold*price-y*cost,
                              'cost':cost,'price':price,'demand':need})
    if not variables:
        return pd.DataFrame(),pd.DataFrame(issues),{'status':'Sin elaboraciones elegibles','investment':0.,'profit':0.}
    n=len(variables)
    if n>2500:
        raise ValueError('Demasiadas elaboraciones candidatas. Reducí el horizonte o la demanda.')
    # budget + batch precedence; second batch cannot be chosen before the first.
    constraints=[np.array([v['investment'] for v in variables],dtype=float)]
    lower=[-np.inf];upper=[float(budget)]
    indices={(v['codigo'],v['batch']):i for i,v in enumerate(variables)}
    for i,v in enumerate(variables):
        if v['batch']:
            r=np.zeros(n);r[i]=1;r[indices[(v['codigo'],v['batch']-1)]]=-1
            constraints.append(r);lower.append(-np.inf);upper.append(0.)
    result=milp(c=-np.array([v['profit'] for v in variables]),integrality=np.ones(n),
                bounds=Bounds(np.zeros(n),np.ones(n)),
                constraints=LinearConstraint(np.vstack(constraints),np.array(lower),np.array(upper)),
                options={'time_limit':20,'mip_rel_gap':0.001})
    if result.x is None:
        raise ValueError('No se encontró una solución del optimizador. Revisá presupuesto y datos.')
    chosen=[v for v,x in zip(variables,result.x) if x>0.5]
    records=[]
    for row in candidates:
        code=str(row['codigo']); group=[v for v in chosen if v['codigo']==code]
        if not group:continue
        batches=len(group); y=group[0]['yield']; need=float(row['demanda_neta'])
        sold=min(need,batches*y);produced=batches*y
        records.append({'CÓDIGO':code,'PRODUCTO':str(row.get('descripcion',code)),
            'ELABORACIONES':batches,'RENDIMIENTO POR ELAB.':y,'BULTOS FABRICADOS':produced,
            'BULTOS VENDIBLES':sold,'EXCEDENTE':produced-sold,
            'DEMANDA SIN CUBRIR':max(0,need-sold),
            'INVERSIÓN':sum(v['investment'] for v in group),
            'VENTA POTENCIAL':sold*group[0]['price'],
            'GANANCIA POTENCIAL':sold*group[0]['price']-sum(v['investment'] for v in group)})
    plan=pd.DataFrame(records)
    return plan,pd.DataFrame(issues),{'status':'Óptimo' if result.status==0 else 'Solución factible (límite de tiempo)',
        'investment':sum(v['investment'] for v in chosen),
        'profit':sum(v['profit'] for v in chosen),
        'selected_batches':len(chosen)}
