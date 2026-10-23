"""Proyecto REP-NFU, grupo 52. Python 3.11 / gurobipy 10.0.3.

Modelo de Complejidad Extrema (Nivel Quincenal):
- Trazabilidad total de la Región de Origen (Multi-commodity).
- Vehículos discretos (Camiones de 10t, 15t y 25t) con variables enteras para viajes.
- Disgregación temporal extrema: 24 Quincenas operativas por año.
- Retorna reportes consolidados anuales exactos (Excel, CSV, resúmenes).

"""

import argparse
import hashlib
import math
import platform
import shutil
import sys
import json
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
import pandas as pd

try:
    import gurobipy as gp
except ImportError:
    gp = None

TIME_LIMIT = 1800.0
MIP_GAP = 0.0001
TOL_TON = 1e-5
ESTADOS = {
    1: "LOADED", 2: "OPTIMAL", 3: "INFEASIBLE", 4: "INF_OR_UNBD",
    5: "UNBOUNDED", 6: "CUTOFF", 7: "ITERATION_LIMIT", 8: "NODE_LIMIT",
    9: "TIME_LIMIT", 10: "SOLUTION_LIMIT", 11: "INTERRUPTED",
    12: "NUMERIC", 13: "SUBOPTIMAL", 14: "INPROGRESS",
    15: "USER_OBJ_LIMIT", 16: "WORK_LIMIT", 17: "MEM_LIMIT",
}

# =============================================================================
# FUNCIONES DE CARGA Y LECTURA
# =============================================================================

def ruta_datos(explicita=None):
    if explicita:
        ruta = Path(explicita).expanduser().resolve()
        if not ruta.is_file():
            raise FileNotFoundError(f"No existe el archivo de entrada: {ruta}")
        return ruta
    base = Path(__file__).resolve().parent
    candidatos = [
        base.parent / "Instancias Excel" / "datos.xlsx",
        base / "Instancias Excel" / "datos.xlsx",
        base.parent / "Instancias_Excel" / "datos.xlsx",
        base / "Instancias_Excel" / "datos.xlsx",
        base / "datos.xlsx", base / "dataset" / "datos.xlsx",
    ]
    for r in candidatos:
        if r.is_file():
            return r
    raise FileNotFoundError("No se encontró datos.xlsx.")

def leer_tabla(libro, hoja, columnas, ids=()):
    df = pd.read_excel(libro, sheet_name=hoja)
    if ids:
        df = df.loc[df[ids[0]].notna()].copy()
        for col in ids:
            df[col] = df[col].astype(str).str.strip()
    return df

def numericos(df, columnas, hoja):
    for col in columnas:
        df[col] = pd.to_numeric(df[col], errors="coerce")

def cargar_datos(ruta):
    libro = pd.ExcelFile(ruta)
    metas = leer_tabla(libro, "Metas_DS8", ["Anio_t", "Meta_Recoleccion_mR [-]", "Meta_Valorizacion_mV [-]", "Cuota_Min_Material_alphaM [-]", "Penalizacion_muR [M$/ton]", "Penalizacion_muV [M$/ton]", "Penalizacion_muQ [M$/ton]"]).set_index("Anio_t").sort_index()
    T = [int(t) for t in metas.index]
    
    regiones = leer_tabla(libro, "Regiones_Generacion", ["ID_Region", "Nombre_Region", "Cuota_DS8_qr [-]"] + [f"Gen_{t} [ton]" for t in T], ["ID_Region"])
    regiones = regiones.loc[regiones.ID_Region.str.upper() != "TOTAL"].set_index("ID_Region")
    
    centros = leer_tabla(libro, "Centros_Acopio", ["ID_Centro", "Nombre_Centro", "Cap_Recepcion_aC [ton/año]", "Cap_Despacho_bC [ton/año]", "Limite_Sanitario_sM [ton]", "Inv_Inicial_s0 [ton]", "Costo_Fijo_Anual_cF [M$/año]", "Costo_Almacenaje_h [M$/ton-año]", "Inversion_Habilitacion_iC [M$]"], ["ID_Centro"]).set_index("ID_Centro")
    
    plantas = leer_tabla(libro, "Plantas_Tecnologias", ["ID_Planta", "Nombre_Planta", "ID_Tecnologia", "Max_Lineas_nP [líneas]", "Factibilidad_epk [-]", "Cap_Efectiva_aP [ton/año]", "Costo_Fijo_cP [M$/año]", "Inversion_Habilitacion_iP [M$]"], ["ID_Planta", "ID_Tecnologia"])
    
    tecnologias = leer_tabla(libro, "Tecnologias_Costos", ["ID_Tecnologia", "Nombre_Tecnologia", "Subconjunto_DS8", "Costo_Variable_cV [M$/ton]"], ["ID_Tecnologia"]).set_index("ID_Tecnologia")
    
    transporte = leer_tabla(libro, "Costos_Transporte", ["Tipo_Tramo", "Nodo_Origen", "Nodo_Destino", "Distancia_Ruta [km]", "Costo_Unitario [M$/ton]"], ["Tipo_Tramo", "Nodo_Origen", "Nodo_Destino"])

    R, C, K = list(regiones.index), list(centros.index), list(tecnologias.index)
    P = sorted(plantas.ID_Planta.unique())
    categorias = tecnologias.Subconjunto_DS8.str.strip()
    KM = list(tecnologias.index[categorias == "Material"])

    pares = plantas.set_index(["ID_Planta", "ID_Tecnologia"])
    d = SimpleNamespace(T=T, R=R, C=C, P=P, K=K, KM=KM, regiones=regiones, centros=centros, plantas=plantas, tecnologias=tecnologias,
        gen={(r, t): float(regiones.loc[r, f"Gen_{t} [ton]"]) for r in R for t in T},
        qr=regiones["Cuota_DS8_qr [-]"].to_dict(),
        mr=metas["Meta_Recoleccion_mR [-]"].to_dict(), mv=metas["Meta_Valorizacion_mV [-]"].to_dict(),
        alpha=0.60, muR=metas["Penalizacion_muR [M$/ton]"].to_dict(),
        muV=metas["Penalizacion_muV [M$/ton]"].to_dict(), muQ=metas["Penalizacion_muQ [M$/ton]"].to_dict(),
        aC=centros["Cap_Recepcion_aC [ton/año]"].to_dict(), bC=centros["Cap_Despacho_bC [ton/año]"].to_dict(),
        sM=centros["Limite_Sanitario_sM [ton]"].to_dict(), s0=centros["Inv_Inicial_s0 [ton]"].to_dict(),
        cF=centros["Costo_Fijo_Anual_cF [M$/año]"].to_dict(), h=centros["Costo_Almacenaje_h [M$/ton-año]"].to_dict(),
        iC=centros["Inversion_Habilitacion_iC [M$]"].to_dict(),
        aP=pares["Cap_Efectiva_aP [ton/año]"].to_dict(), epk=pares["Factibilidad_epk [-]"].to_dict(),
        cP=pares["Costo_Fijo_cP [M$/año]"].to_dict(), iP=pares["Inversion_Habilitacion_iP [M$]"].to_dict(),
        nP=plantas.groupby("ID_Planta")["Max_Lineas_nP [líneas]"].first().to_dict(),
        cV=tecnologias["Costo_Variable_cV [M$/ton]"].to_dict())

    for tipo, nombre in [("T1_Region_Centro", "T1"), ("T2_Centro_Planta", "T2"), ("T3_Region_Planta", "T3")]:
        tramo = transporte.loc[transporte.Tipo_Tramo == tipo].set_index(["Nodo_Origen", "Nodo_Destino"])
        setattr(d, "c" + nombre, tramo["Costo_Unitario [M$/ton]"].to_dict())
        setattr(d, "dist" + nombre, tramo["Distancia_Ruta [km]"].to_dict())
        
    d.origen_inversion = "Montos ingresados en Excel"
    if "Supuestos_Inversion" in libro.sheet_names:
        supuestos = pd.read_excel(libro, sheet_name="Supuestos_Inversion").set_index("Parametro")
        d.origen_inversion = str(supuestos.loc["Origen_Inversiones", "Valor"])
    libro.close()
    return d

# =============================================================================
# MODELO MATEMÁTICO 
# =============================================================================

def construir_modelo_extremo(d, salida):
    GRB, Q = gp.GRB, gp.quicksum
    m = gp.Model("Optimizacion_REP_NFU_Extremo")
    m.Params.LogFile = str(salida / "gurobi.log")
    m.Params.TimeLimit = TIME_LIMIT
    m.Params.MIPGap = MIP_GAP
    m.Params.Seed = 52
    
    VEHICULOS = {
        'Camion_10t': {'cap': 10, 'costo_f': 1.2},
        'Camion_15t': {'cap': 15, 'costo_f': 1.1},
        'Camion_25t': {'cap': 25, 'costo_f': 1.0}
    }
    V = list(VEHICULOS.keys())
    Qu = list(range(1, 25))
    
    d.V = V
    d.Qu = Qu
    
    y = m.addVars(d.C, d.T, vtype=GRB.BINARY, name="y")
    w = m.addVars(d.P, d.K, d.T, vtype=GRB.BINARY, name="w")
    
    nx = m.addVars(d.R, d.C, V, d.T, Qu, vtype=GRB.INTEGER, lb=0, name="nx")
    nz = m.addVars(d.R, d.P, V, d.T, Qu, vtype=GRB.INTEGER, lb=0, name="nz")
    nf = m.addVars(d.R, d.C, d.P, V, d.T, Qu, vtype=GRB.INTEGER, lb=0, name="nf")
    
    x = m.addVars(d.R, d.C, V, d.T, Qu, lb=0, name="x")
    z = m.addVars(d.R, d.P, V, d.T, Qu, lb=0, name="z")
    f = m.addVars(d.R, d.C, d.P, V, d.T, Qu, lb=0, name="f")
    v = m.addVars(d.R, d.P, d.K, d.T, Qu, lb=0, name="v")
    s = m.addVars(d.R, d.C, d.T, Qu, lb=0, name="s")
    
    dR = m.addVars(d.T, Qu, lb=0, name="dR")
    dV = m.addVars(d.T, Qu, lb=0, name="dV")
    dQ = m.addVars(d.R, d.T, lb=0, name="dQ")

    anterior = {t: d.T[i-1] if i else None for i, t in enumerate(d.T)}
    dy = {(c, t): y[c, t] - (y[c, anterior[t]] if anterior[t] is not None else 0) for c in d.C for t in d.T}
    dw = {(p, k, t): w[p, k, t] - (w[p, k, anterior[t]] if anterior[t] is not None else 0) for p in d.P for k in d.K for t in d.T}
    
    costos = {
        "IC_C inversion habilitacion centros": Q(d.iC[c] * dy[c, t] for c in d.C for t in d.T),
        "IC_P inversion habilitacion lineas": Q(d.iP[p, k] * dw[p, k, t] for p in d.P for k in d.K for t in d.T),
        "CF_C operacion fija centros": Q(d.cF[c] * y[c, t] for c in d.C for t in d.T),
        "CF_P operacion fija plantas": Q(d.cP[p, k] * w[p, k, t] for p in d.P for k in d.K for t in d.T),
        "CT_1 transporte region-acopio": Q(d.cT1[o, c] * VEHICULOS[vh]['costo_f'] * x[o, c, vh, t, qu] for o in d.R for c in d.C for vh in V for t in d.T for qu in Qu),
        "CT_2 transporte acopio-planta": Q(d.cT2[c, p] * VEHICULOS[vh]['costo_f'] * f[o, c, p, vh, t, qu] for o in d.R for c in d.C for p in d.P for vh in V for t in d.T for qu in Qu),
        "CT_3 transporte region-planta": Q(d.cT3[o, p] * VEHICULOS[vh]['costo_f'] * z[o, p, vh, t, qu] for o in d.R for p in d.P for vh in V for t in d.T for qu in Qu),
        "CP_V procesamiento": Q(d.cV[k] * v[o, p, k, t, qu] for o in d.R for p in d.P for k in d.K for t in d.T for qu in Qu),
        "CA_S almacenaje": Q((d.h[c]/24) * s[o, c, t, qu] for o in d.R for c in d.C for t in d.T for qu in Qu),
        "CP_M penalizaciones deficit": Q(d.muR[t] * dR[t, qu] + d.muV[t] * dV[t, qu] for t in d.T for qu in Qu) + Q(d.muQ[t] * dQ[o, t] for o in d.R for t in d.T),
    }
    
    m.setObjective(Q(costos.values()), GRB.MINIMIZE)

    m.addConstrs((x[o, c, vh, t, qu] <= VEHICULOS[vh]['cap'] * nx[o, c, vh, t, qu] for o in d.R for c in d.C for vh in V for t in d.T for qu in Qu), "R_RutaX")
    m.addConstrs((z[o, p, vh, t, qu] <= VEHICULOS[vh]['cap'] * nz[o, p, vh, t, qu] for o in d.R for p in d.P for vh in V for t in d.T for qu in Qu), "R_RutaZ")
    m.addConstrs((f[o, c, p, vh, t, qu] <= VEHICULOS[vh]['cap'] * nf[o, c, p, vh, t, qu] for o in d.R for c in d.C for p in d.P for vh in V for t in d.T for qu in Qu), "R_RutaF")
    
    m.addConstrs((Q(x[r, c, vh, t, qu] for c in d.C for vh in V) + Q(z[r, p, vh, t, qu] for p in d.P for vh in V) <= d.gen[r, t]/24 for r in d.R for t in d.T for qu in Qu), "R1")
    
    m.addConstrs((Q(x[o, c, vh, t, qu] for o in d.R for vh in V) <= (d.aC[c]/24) * y[c, t] for c in d.C for t in d.T for qu in Qu), "R2a")
    m.addConstrs((Q(f[o, c, p, vh, t, qu] for o in d.R for p in d.P for vh in V) <= (d.bC[c]/24) * y[c, t] for c in d.C for t in d.T for qu in Qu), "R2b")
    m.addConstrs((Q(s[o, c, t, qu] for o in d.R) <= d.sM[c] * y[c, t] for c in d.C for t in d.T for qu in Qu), "R3")
    
    for o in d.R:
        for c in d.C:
            for t in d.T:
                for qu in Qu:
                    if qu > 1:
                        previo = s[o, c, t, qu-1]
                    else:
                        previo = s[o, c, anterior[t], 24] if anterior[t] is not None else 0
                    m.addConstr(s[o, c, t, qu] == previo + Q(x[o, c, vh, t, qu] for vh in V) - Q(f[o, c, p, vh, t, qu] for p in d.P for vh in V), f"R4_{o}_{c}_{t}_{qu}")
    
    m.addConstrs((Q(v[o, p, k, t, qu] for k in d.K) == Q(f[o, c, p, vh, t, qu] for c in d.C for vh in V) + Q(z[o, p, vh, t, qu] for vh in V) for o in d.R for p in d.P for t in d.T for qu in Qu), "R5")
    
    m.addConstrs((Q(v[o, p, k, t, qu] for o in d.R) <= (d.aP[p, k]/24) * w[p, k, t] for p in d.P for k in d.K for t in d.T for qu in Qu), "R6")
    m.addConstrs((w[p, k, t] <= d.epk[p, k] for p in d.P for k in d.K for t in d.T), "R6b")
    m.addConstrs((Q(w[p, k, t] for k in d.K) <= d.nP[p] for p in d.P for t in d.T), "R6c")
    
    for t in d.T:
        gen_qu_total = sum(d.gen[r, t] for r in d.R) / 24
        for qu in Qu:
            m.addConstr(Q(x[o, c, vh, t, qu] for o in d.R for c in d.C for vh in V) + Q(z[o, p, vh, t, qu] for o in d.R for p in d.P for vh in V) + dR[t, qu] >= d.mr[t] * gen_qu_total, "R7")
            m.addConstr(Q(v[o, p, k, t, qu] for o in d.R for p in d.P for k in d.K) + dV[t, qu] >= d.mv[t] * gen_qu_total, "R8")
            m.addConstr(Q(v[o, p, k, t, qu] for o in d.R for p in d.P for k in d.KM) >= d.alpha * Q(v[o, p, k, t, qu] for o in d.R for p in d.P for k in d.K), "R12")
        for r in d.R:
            m.addConstr(Q(x[r, c, vh, t, qu] for c in d.C for vh in V for qu in Qu) + Q(z[r, p, vh, t, qu] for p in d.P for vh in V for qu in Qu) + dQ[r, t] >= d.qr[r] * sum(d.gen[o, t] for o in d.R), "R9")
            
    for t in d.T[1:]:
        m.addConstrs((y[c, t] >= y[c, anterior[t]] for c in d.C), "R10a")
        m.addConstrs((w[p, k, t] >= w[p, k, anterior[t]] for p in d.P for k in d.K), "R10b")
        
    m.update()
    
    variables = dict(y=y, w=w, x=x, z=z, f=f, v=v, s=s, dR=dR, dV=dV, dQ=dQ, nx=nx, nz=nz, nf=nf)
    return m, variables, costos, anterior

# =============================================================================
# EXTRACCIÓN Y CONSOLIDACIÓN DE RESULTADOS
# =============================================================================

def finito(valor):
    if valor is None: return None
    valor = float(valor)
    return valor if math.isfinite(valor) and abs(valor) < 1e90 else None

def atributo(modelo, nombre):
    try: return finito(modelo.getAttr(nombre))
    except: return None

def progreso_callback(historial):
    ultimo = [-5.0]
    def registrar(modelo, donde):
        if donde != gp.GRB.Callback.MIP: return
        tiempo = modelo.cbGet(gp.GRB.Callback.RUNTIME)
        if tiempo - ultimo[0] < 5.0: return
        ultimo[0] = tiempo
        obj = finito(modelo.cbGet(gp.GRB.Callback.MIP_OBJBST))
        cota = finito(modelo.cbGet(gp.GRB.Callback.MIP_OBJBND))
        gap = abs(obj-cota) / abs(obj) if obj and cota and obj != 0 else None
        historial.append(dict(tiempo_s=tiempo, objetivo_M_CLP=obj, cota_M_CLP=cota, gap_pct=100*gap if gap else None, nodos=modelo.cbGet(gp.GRB.Callback.MIP_NODCNT)))
    return registrar

def guardar_resultados(m, d, var, costos, anterior, salida, ruta, historial):
    hay_solucion = m.SolCount > 0
    obj = atributo(m, "ObjVal") if hay_solucion else None
    cota = atributo(m, "ObjBound")
    gap = atributo(m, "MIPGap") if hay_solucion else None
    
    info = dict(estado_codigo=m.Status, estado=ESTADOS.get(m.Status, str(m.Status)),
        hay_solucion=hay_solucion, optimalidad_certificada_en_tolerancia=m.Status == gp.GRB.OPTIMAL,
        objetivo_M_CLP=obj, objetivo_CLP=obj*1000 if obj is not None else None,
        cota_M_CLP=cota, gap_pct=100*gap if gap is not None else None,
        tiempo_solver_s=m.Runtime, limite_s=TIME_LIMIT,
        nodos=atributo(m, "NodeCount"), variables=m.NumVars,
        binarias=m.NumBinVars, enteras=m.NumIntVars - m.NumBinVars, restricciones=m.NumConstrs)
    
    if len(historial) == 0 or historial[-1].get("tiempo_s", 0) != m.Runtime:
        historial.append(dict(tiempo_s=m.Runtime, objetivo_M_CLP=obj, cota_M_CLP=cota, gap_pct=info.get("gap_pct"), nodos=info.get("nodos")))
        
    tablas = {"Historial_Gurobi": pd.DataFrame(historial)}
    resumen = [f"Estado: {info['estado']} (código {m.Status}).",
               f"Tiempo de optimización: {m.Runtime:.3f} s; límite: {TIME_LIMIT:.0f} s.",
               f"Variables Totales: {m.NumVars} (Binarias: {m.NumBinVars}, Enteras: {m.NumIntVars - m.NumBinVars}, Continuas: {m.NumVars - m.NumIntVars})",
               f"Restricciones: {m.NumConstrs}; soluciones almacenadas: {m.SolCount}."]
               
    if hay_solucion:
        def X(nombre, *indices):
            clave = indices[0] if len(indices) == 1 else tuple(indices)
            return float(var[nombre][clave].X)
            
        costos_val = {nombre: float(exp.getValue()) for nombre, exp in costos.items()}
        total = sum(costos_val.values())
        filas_costos = [dict(Componente=k, Costo_M_CLP=v, Costo_CLP=1000*v, Pct=100*v/total if total != 0 else None) for k, v in costos_val.items()]
        filas_costos.append(dict(Componente="TOTAL", Costo_M_CLP=total, Costo_CLP=1000*total, Pct=100 if total else None))
        tablas["Resumen_Costos_FO"] = pd.DataFrame(filas_costos)
        
        ds8, cuotas, acopios, plantas, flujos, aperturas, balances = [], [], [], [], [], [], []
        nr = d.regiones.Nombre_Region.to_dict()
        nc = d.centros.Nombre_Centro.to_dict()
        np_ = d.plantas.groupby("ID_Planta").Nombre_Planta.first().to_dict()
        nk = d.tecnologias.Nombre_Tecnologia.to_dict()
        
        for t in d.T:
            gen = sum(d.gen[r, t] for r in d.R)
            
            # Acumulando dimensiones: origenes, vehiculos, quincenas
            rec = sum(X("x", o, c, vh, t, qu) for o in d.R for c in d.C for vh in d.V for qu in d.Qu) + sum(X("z", o, p, vh, t, qu) for o in d.R for p in d.P for vh in d.V for qu in d.Qu)
            val = sum(X("v", o, p, k, t, qu) for o in d.R for p in d.P for k in d.K for qu in d.Qu)
            mat = sum(X("v", o, p, k, t, qu) for o in d.R for p in d.P for k in d.KM for qu in d.Qu)
            
            # Stock a fin de año (quincena 24) y stock previo (año anterior, quincena 24)
            stock_previo = sum(X("s", o, c, anterior[t], 24) if anterior[t] is not None else 0 for o in d.R for c in d.C)
            stock_final = sum(X("s", o, c, t, 24) for o in d.R for c in d.C)
            residuo = rec + stock_previo - val - stock_final
            
            balances.append(dict(anio=t, recoleccion_ton=rec, inventario_inicial_ton=stock_previo, valorizacion_ton=val, inventario_final_ton=stock_final, residuo_balance_ton=residuo))
            
            ds8.append(dict(anio=t, generacion_ton=gen, meta_rec_ton=d.mr[t]*gen, recoleccion_ton=rec, pct_recoleccion=100*rec/gen, 
                deficit_rec_ton=sum(X("dR", t, qu) for qu in d.Qu), meta_val_ton=d.mv[t]*gen, valorizacion_ton=val, pct_valorizacion=100*val/gen, 
                deficit_val_ton=sum(X("dV", t, qu) for qu in d.Qu), valorizacion_material_ton=mat, valorizacion_energetica_ton=val-mat,
                pct_material_sobre_valorizado=100*mat/val if val > TOL_TON else None, minimo_material_ton=d.alpha*val, holgura_material_ton=mat-d.alpha*val,
                deficit_cuotas_regionales_ton=sum(X("dQ", r, t) for r in d.R)))
                
            for r in d.R:
                recogido = sum(X("x", r, c, vh, t, qu) for c in d.C for vh in d.V for qu in d.Qu) + sum(X("z", r, p, vh, t, qu) for p in d.P for vh in d.V for qu in d.Qu)
                cuotas.append(dict(anio=t, region_id=r, region=nr[r], disponible_ton=d.gen[r, t], cuota_minima_ton=d.qr[r]*gen, recoleccion_ton=recogido, deficit_ton=X("dQ", r, t)))
                
            for c in d.C:
                activo = X("y", c, t)
                apertura = activo - (X("y", c, anterior[t]) if anterior[t] is not None else 0)
                entrada = sum(X("x", o, c, vh, t, qu) for o in d.R for vh in d.V for qu in d.Qu)
                despacho = sum(X("f", o, c, p, vh, t, qu) for o in d.R for p in d.P for vh in d.V for qu in d.Qu)
                acopios.append(dict(anio=t, centro_id=c, centro=nc[c], activo=round(activo), cap_entrada_ton=d.aC[c]*activo, flujo_entrada_ton=entrada,
                    uso_entrada_pct=100*entrada/(d.aC[c]*activo) if d.aC[c]*activo > TOL_TON else None, cap_salida_ton=d.bC[c]*activo, flujo_salida_ton=despacho,
                    uso_salida_pct=100*despacho/(d.bC[c]*activo) if d.bC[c]*activo > TOL_TON else None, inv_max_ton=d.sM[c]*activo, 
                    inventario_final_ton=sum(X("s", o, c, t, 24) for o in d.R)))
                aperturas.append(dict(anio=t, tipo="Centro", instalacion=c, tecnologia="", activo=round(activo), apertura=round(apertura), inversion_M_CLP=d.iC[c]*apertura))
                
            for p in d.P:
                for k in d.K:
                    activo = X("w", p, k, t)
                    apertura = activo - (X("w", p, k, anterior[t]) if anterior[t] is not None else 0)
                    proc = sum(X("v", o, p, k, t, qu) for o in d.R for qu in d.Qu)
                    plantas.append(dict(anio=t, planta_id=p, planta=np_[p], tecnologia_id=k, tecnologia=nk[k], categoria_ds8="Material" if k in d.KM else "Energetica", 
                        compatible=d.epk[p, k], activa=round(activo), cap_efectiva_ton=d.aP[p, k]*activo, valorizado_ton=proc,
                        uso_capacidad_pct=100*proc/(d.aP[p, k]*activo) if d.aP[p, k]*activo > TOL_TON else None))
                    aperturas.append(dict(anio=t, tipo="Linea de planta", instalacion=p, tecnologia=k, activo=round(activo), apertura=round(apertura), inversion_M_CLP=d.iP[p, k]*apertura))
            
            # Flujos anuales consolidados
            for r in d.R:
                for c in d.C:
                    flujo = sum(X("x", r, c, vh, t, qu) for vh in d.V for qu in d.Qu)
                    if flujo > TOL_TON:
                        flujos.append(dict(anio=t, tramo="T1_Region_Acopio", origen_id=r, origen=nr[r], destino_id=c, destino=nc[c], distancia_km=d.distT1[r, c], costo_unit_M_CLP=d.cT1[r, c], toneladas=flujo))
            for c in d.C:
                for p in d.P:
                    flujo = sum(X("f", o, c, p, vh, t, qu) for o in d.R for vh in d.V for qu in d.Qu)
                    if flujo > TOL_TON:
                        flujos.append(dict(anio=t, tramo="T2_Acopio_Planta", origen_id=c, origen=nc[c], destino_id=p, destino=np_[p], distancia_km=d.distT2[c, p], costo_unit_M_CLP=d.cT2[c, p], toneladas=flujo))
            for r in d.R:
                for p in d.P:
                    flujo = sum(X("z", r, p, vh, t, qu) for vh in d.V for qu in d.Qu)
                    if flujo > TOL_TON:
                        flujos.append(dict(anio=t, tramo="T3_Region_Planta_Directo", origen_id=r, origen=nr[r], destino_id=p, destino=np_[p], distancia_km=d.distT3[r, p], costo_unit_M_CLP=d.cT3[r, p], toneladas=flujo))
                        
        tablas.update({"Cumplimiento_Normativo_DS8":pd.DataFrame(ds8), "Cuotas_Regionales":pd.DataFrame(cuotas),
            "Capacidad_Centros_Acopio":pd.DataFrame(acopios), "Capacidad_Plantas_Tec":pd.DataFrame(plantas),
            "Matriz_Flujos_Transporte":pd.DataFrame(flujos, columns=["anio","tramo","origen_id","origen","destino_id","destino","distancia_km","costo_unit_M_CLP","toneladas"]),
            "Aperturas_Inversion":pd.DataFrame(aperturas), "Balance_Global":pd.DataFrame(balances)})
            
        m.write(str(salida / "solucion.sol"))
    
    estado = pd.DataFrame([info])
    estado.to_csv(salida / "resumen_solver.csv", index=False, encoding="utf-8-sig")
    
    with pd.ExcelWriter(salida / "Resultados_Optimizacion_REP_NFU_Extremo.xlsx") as writer:
        estado.to_excel(writer, sheet_name="Estado_Gurobi", index=False)
        for hoja, df in tablas.items():
            df.to_excel(writer, sheet_name=hoja, index=False)
            
    texto = "\n".join(resumen) + "\n"
    (salida / "resumen_ejecucion.txt").write_text(texto, encoding="utf-8")
    print("\n" + texto)
    print(f"Archivos guardados en: {salida}")
    return info

# =============================================================================
# MAIN EXECUTOR
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description="Red REP-NFU Modelo de Complejidad Extrema")
    args = parser.parse_args()
    
    ruta = ruta_datos()
    d = cargar_datos(ruta)
    print(f"Excel cargado. {len(d.R)} regiones; {len(d.C)} centros; {len(d.P)} plantas; {len(d.K)} tecnologías; años {d.T}.")
    
    base = Path(__file__).resolve().parent
    raiz = base.parent if base.name in ["Script Python", "Script_Python"] else base
    carpeta = raiz / "Analisis de Excels"
    salida = carpeta / datetime.now().astimezone().strftime("ejecucion_extrema_%Y%m%d_%H%M%S_%f")
    salida.mkdir(parents=True, exist_ok=True)
    shutil.copy2(ruta, salida / "entrada_datos.xlsx")
    
    m = None
    try:
        m, variables, costos, anterior = construir_modelo_extremo(d, salida)
        print(f"\n--- ESTADÍSTICAS DEL MODELO DE COMPLEJIDAD EXTREMA ---")
        print(f"Variables totales: {m.NumVars:,}")
        print(f"  Binarias: {m.NumBinVars:,}")
        print(f"  Enteras: {m.NumIntVars - m.NumBinVars:,}")
        print(f"  Continuas: {m.NumVars - m.NumIntVars:,}")
        print(f"Restricciones totales: {m.NumConstrs:,}")
        
        historial = []
        m.optimize(progreso_callback(historial))
        info = guardar_resultados(m, d, variables, costos, anterior, salida, ruta, historial)
        
        return 0 if info["hay_solucion"] else 2
    finally:
        if m is not None:
            m.dispose()

if __name__ == "__main__":
    sys.exit(main())