"""Proyecto REP-NFU, grupo 52. Python 3.11 / gurobipy 10.0.3.

Ejecutar: python main.py
Validar Excel sin resolver: python main.py --validar-datos
Unidades de costos: M$ = miles de pesos chilenos; MM$ = millones de pesos.
Las inversiones de la instancia adjunta son supuestos de PRUEBA editables.
"""

import argparse
import hashlib
import math
import platform
import shutil
import sys
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

import pandas as pd

try:
    import gurobipy as gp
except ImportError:
    gp = None  # --validar-datos no necesita Gurobi ni una licencia.

TIME_LIMIT = 1800.0  # E3: 30 minutos de optimización.
MIP_GAP = 0.0001  # Fracción: 0,0001 equivale a 0,01 %, no a 0,0001 %.
TOL_TON = 1e-5
ESTADOS = {
    1: "LOADED", 2: "OPTIMAL", 3: "INFEASIBLE", 4: "INF_OR_UNBD",
    5: "UNBOUNDED", 6: "CUTOFF", 7: "ITERATION_LIMIT", 8: "NODE_LIMIT",
    9: "TIME_LIMIT", 10: "SOLUTION_LIMIT", 11: "INTERRUPTED",
    12: "NUMERIC", 13: "SUBOPTIMAL", 14: "INPROGRESS",
    15: "USER_OBJ_LIMIT", 16: "WORK_LIMIT", 17: "MEM_LIMIT",
}


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
    for ruta in candidatos:
        if ruta.is_file():
            return ruta
    raise FileNotFoundError("No se encontró datos.xlsx. Ubícalo en Instancias Excel/ "
                            "o usa --datos RUTA. La búsqueda usa la carpeta del script.")


def leer_tabla(libro, hoja, columnas, ids=()):
    df = pd.read_excel(libro, sheet_name=hoja)
    faltan = sorted(set(columnas) - set(df.columns))
    if faltan:
        raise ValueError(f"{hoja}: faltan columnas {faltan}")
    if ids:
        df = df.loc[df[ids[0]].notna()].copy()
        for col in ids:
            if df[col].isna().any():
                raise ValueError(f"{hoja}: identificadores vacíos en {col}")
            df[col] = df[col].astype(str).str.strip()
        if df.duplicated(list(ids)).any():
            raise ValueError(f"{hoja}: claves duplicadas en {ids}")
    return df


def numericos(df, columnas, hoja):
    for col in columnas:
        df[col] = pd.to_numeric(df[col], errors="raise")
        if not df[col].map(lambda x: pd.notna(x) and math.isfinite(x) and x >= 0).all():
            raise ValueError(f"{hoja}: {col} debe contener números no negativos, sin vacíos.")


def cargar_datos(ruta):
    """Lee parámetros; no supone que un dato faltante sea cero."""
    libro = pd.ExcelFile(ruta)
    metas = leer_tabla(libro, "Metas_DS8", ["Anio_t", "Meta_Recoleccion_mR [-]",
        "Meta_Valorizacion_mV [-]", "Cuota_Min_Material_alphaM [-]",
        "Penalizacion_muR [M$/ton]", "Penalizacion_muV [M$/ton]", "Penalizacion_muQ [M$/ton]"])
    numericos(metas, metas.columns, "Metas_DS8")
    if metas.empty or metas.Anio_t.duplicated().any() or (metas.Anio_t % 1 != 0).any():
        raise ValueError("Metas_DS8: los años deben ser enteros, únicos y no vacíos.")
    metas = metas.set_index("Anio_t").sort_index()
    T = [int(t) for t in metas.index]
    if T != list(range(T[0], T[-1] + 1)):
        raise ValueError("Metas_DS8: el horizonte debe contener años consecutivos.")
    for col in ["Meta_Recoleccion_mR [-]", "Meta_Valorizacion_mV [-]", "Cuota_Min_Material_alphaM [-]"]:
        if (metas[col] > 1).any():
            raise ValueError(f"Metas_DS8: {col} debe estar entre 0 y 1.")
    if not (metas["Cuota_Min_Material_alphaM [-]"].sub(0.60).abs() < 1e-10).all():
        raise ValueError("La instancia base debe exigir alpha = 0,60 en todos los años.")

    regiones = leer_tabla(libro, "Regiones_Generacion",
        ["ID_Region", "Nombre_Region", "Cuota_DS8_qr [-]"] + [f"Gen_{t} [ton]" for t in T], ["ID_Region"])
    regiones = regiones.loc[regiones.ID_Region.str.upper() != "TOTAL"].set_index("ID_Region")
    centros = leer_tabla(libro, "Centros_Acopio", ["ID_Centro", "Nombre_Centro",
        "Cap_Recepcion_aC [ton/año]", "Cap_Despacho_bC [ton/año]", "Limite_Sanitario_sM [ton]",
        "Inv_Inicial_s0 [ton]", "Costo_Fijo_Anual_cF [M$/año]", "Costo_Almacenaje_h [M$/ton-año]",
        "Inversion_Habilitacion_iC [M$]"], ["ID_Centro"]).set_index("ID_Centro")
    plantas = leer_tabla(libro, "Plantas_Tecnologias", ["ID_Planta", "Nombre_Planta", "ID_Tecnologia",
        "Max_Lineas_nP [líneas]", "Factibilidad_epk [-]", "Cap_Efectiva_aP [ton/año]",
        "Costo_Fijo_cP [M$/año]", "Inversion_Habilitacion_iP [M$]"], ["ID_Planta", "ID_Tecnologia"])
    tecnologias = leer_tabla(libro, "Tecnologias_Costos", ["ID_Tecnologia", "Nombre_Tecnologia",
        "Subconjunto_DS8", "Costo_Variable_cV [M$/ton]"], ["ID_Tecnologia"]).set_index("ID_Tecnologia")
    transporte = leer_tabla(libro, "Costos_Transporte", ["Tipo_Tramo", "Nodo_Origen", "Nodo_Destino",
        "Distancia_Ruta [km]", "Costo_Unitario [M$/ton]"], ["Tipo_Tramo", "Nodo_Origen", "Nodo_Destino"])

    numericos(regiones, ["Cuota_DS8_qr [-]"] + [f"Gen_{t} [ton]" for t in T], "Regiones_Generacion")
    numericos(centros, [c for c in centros.columns if "[" in c], "Centros_Acopio")
    numericos(plantas, ["Max_Lineas_nP [líneas]", "Factibilidad_epk [-]", "Cap_Efectiva_aP [ton/año]",
        "Costo_Fijo_cP [M$/año]", "Inversion_Habilitacion_iP [M$]"], "Plantas_Tecnologias")
    numericos(tecnologias, ["Costo_Variable_cV [M$/ton]"], "Tecnologias_Costos")
    numericos(transporte, ["Distancia_Ruta [km]", "Costo_Unitario [M$/ton]"], "Costos_Transporte")
    R, C, K = list(regiones.index), list(centros.index), list(tecnologias.index)
    P = sorted(plantas.ID_Planta.unique())
    if not all([R, C, K, P]):
        raise ValueError("Los conjuntos de regiones, centros, plantas y tecnologías deben ser no vacíos.")
    if set(zip(plantas.ID_Planta, plantas.ID_Tecnologia)) != {(p, k) for p in P for k in K}:
        raise ValueError("Plantas_Tecnologias debe contener cada par planta-tecnología una sola vez.")
    if not plantas["Factibilidad_epk [-]"].isin([0, 1]).all():
        raise ValueError("Factibilidad_epk debe ser binaria.")
    if (plantas["Max_Lineas_nP [líneas]"] % 1 != 0).any():
        raise ValueError("Max_Lineas_nP debe ser entero.")
    if (plantas.groupby("ID_Planta")["Max_Lineas_nP [líneas]"].nunique() != 1).any():
        raise ValueError("Una planta debe tener el mismo máximo de líneas en todas sus filas.")
    if (centros["Inv_Inicial_s0 [ton]"] != 0).any():
        raise ValueError("El informe supone una red nueva con inventario inicial cero.")
    if (regiones["Cuota_DS8_qr [-]"] > 1).any():
        raise ValueError("Las cuotas regionales deben estar entre 0 y 1.")
    categorias = tecnologias.Subconjunto_DS8.str.strip()
    if not categorias.isin(["Material", "Energetica"]).all():
        raise ValueError("Subconjunto_DS8 debe identificar Material o Energetica para cada tecnología.")
    KM = list(tecnologias.index[categorias == "Material"])
    if not KM:
        raise ValueError("Faltan tecnologías de reciclaje material o recauchaje.")

    # Diccionarios con la misma notación operativa del código original.
    pares = plantas.set_index(["ID_Planta", "ID_Tecnologia"])
    d = SimpleNamespace(T=T, R=R, C=C, P=P, K=K, KM=KM, regiones=regiones, centros=centros,
        plantas=plantas, tecnologias=tecnologias,
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
    for t in T:
        if sum(d.gen[r, t] for r in R) <= 0:
            raise ValueError(f"Generación nacional nula en {t}.")
    for tipo, origenes, destinos, nombre in [
        ("T1_Region_Centro", R, C, "T1"), ("T2_Centro_Planta", C, P, "T2"), ("T3_Region_Planta", R, P, "T3")]:
        tramo = transporte.loc[transporte.Tipo_Tramo == tipo].set_index(["Nodo_Origen", "Nodo_Destino"])
        if set(tramo.index) != {(o, q) for o in origenes for q in destinos}:
            raise ValueError(f"Costos_Transporte: faltan o sobran arcos para {tipo}.")
        setattr(d, "c" + nombre, tramo["Costo_Unitario [M$/ton]"].to_dict())
        setattr(d, "dist" + nombre, tramo["Distancia_Ruta [km]"].to_dict())
    d.origen_inversion = "Montos ingresados en Excel"
    if "Supuestos_Inversion" in libro.sheet_names:
        supuestos = pd.read_excel(libro, sheet_name="Supuestos_Inversion").set_index("Parametro")
        d.origen_inversion = str(supuestos.loc["Origen_Inversiones", "Valor"])
    libro.close()
    return d


def construir_modelo(d, salida):
    GRB, Q = gp.GRB, gp.quicksum
    m = gp.Model("Optimizacion_REP_NFU")
    m.Params.LogFile = str(salida / "gurobi.log")
    m.Params.TimeLimit = TIME_LIMIT
    m.Params.MIPGap = MIP_GAP
    m.Params.Seed = 52
    y = m.addVars(d.C, d.T, vtype=GRB.BINARY, name="y")
    w = m.addVars(d.P, d.K, d.T, vtype=GRB.BINARY, name="w")
    x = m.addVars(d.R, d.C, d.T, lb=0, name="x")
    z = m.addVars(d.R, d.P, d.T, lb=0, name="z")
    f = m.addVars(d.C, d.P, d.T, lb=0, name="f")
    v = m.addVars(d.P, d.K, d.T, lb=0, name="v")
    s = m.addVars(d.C, d.T, lb=0, name="s")
    dR = m.addVars(d.T, lb=0, name="dR")
    dV = m.addVars(d.T, lb=0, name="dV")
    dQ = m.addVars(d.R, d.T, lb=0, name="dQ")
    anterior = {t: d.T[i-1] if i else None for i, t in enumerate(d.T)}
    dy = {(c, t): y[c, t] - (y[c, anterior[t]] if anterior[t] is not None else 0) for c in d.C for t in d.T}
    dw = {(p, k, t): w[p, k, t] - (w[p, k, anterior[t]] if anterior[t] is not None else 0)
          for p in d.P for k in d.K for t in d.T}
    costos = {
        "IC_C inversion habilitacion centros": Q(d.iC[c] * dy[c, t] for c in d.C for t in d.T),
        "IC_P inversion habilitacion lineas": Q(d.iP[p, k] * dw[p, k, t] for p in d.P for k in d.K for t in d.T),
        "CF_C operacion fija centros": Q(d.cF[c] * y[c, t] for c in d.C for t in d.T),
        "CF_P operacion fija plantas": Q(d.cP[p, k] * w[p, k, t] for p in d.P for k in d.K for t in d.T),
        "CT_1 transporte region-acopio": Q(d.cT1[r, c] * x[r, c, t] for r in d.R for c in d.C for t in d.T),
        "CT_2 transporte acopio-planta": Q(d.cT2[c, p] * f[c, p, t] for c in d.C for p in d.P for t in d.T),
        "CT_3 transporte region-planta": Q(d.cT3[r, p] * z[r, p, t] for r in d.R for p in d.P for t in d.T),
        "CP_V procesamiento": Q(d.cV[k] * v[p, k, t] for p in d.P for k in d.K for t in d.T),
        "CA_S almacenaje": Q(d.h[c] * s[c, t] for c in d.C for t in d.T),
        "CP_M penalizaciones deficit": Q(d.muR[t] * dR[t] + d.muV[t] * dV[t] + Q(d.muQ[t] * dQ[r, t] for r in d.R) for t in d.T),
    }
    m.setObjective(Q(costos.values()), GRB.MINIMIZE)
    # R1: disponibilidad regional.
    m.addConstrs((Q(x[r, c, t] for c in d.C) + Q(z[r, p, t] for p in d.P) <= d.gen[r, t]
                  for r in d.R for t in d.T), name="R1_disponibilidad")
    # R2a y R2b: entrada y salida, conservando las dos capacidades del Excel.
    m.addConstrs((Q(x[r, c, t] for r in d.R) <= d.aC[c] * y[c, t] for c in d.C for t in d.T), name="R2a_recepcion")
    m.addConstrs((Q(f[c, p, t] for p in d.P) <= d.bC[c] * y[c, t] for c in d.C for t in d.T), name="R2b_despacho")
    m.addConstrs((s[c, t] <= d.sM[c] * y[c, t] for c in d.C for t in d.T), name="R3_stock")
    for c in d.C:
        for t in d.T:
            previo = s[c, anterior[t]] if anterior[t] is not None else d.s0[c]
            m.addConstr(s[c, t] == previo + Q(x[r, c, t] for r in d.R) - Q(f[c, p, t] for p in d.P), name=f"R4_balance[{c},{t}]")
    m.addConstrs((Q(v[p, k, t] for k in d.K) == Q(f[c, p, t] for c in d.C) + Q(z[r, p, t] for r in d.R)
                  for p in d.P for t in d.T), name="R5_planta_sin_stock")
    m.addConstrs((v[p, k, t] <= d.aP[p, k] * w[p, k, t] for p in d.P for k in d.K for t in d.T), name="R6_capacidad")
    m.addConstrs((w[p, k, t] <= d.epk[p, k] for p in d.P for k in d.K for t in d.T), name="R6b_compatibilidad")
    m.addConstrs((Q(w[p, k, t] for k in d.K) <= d.nP[p] for p in d.P for t in d.T), name="R6c_max_lineas")
    m.addConstrs((Q(x[r, c, t] for r in d.R for c in d.C) + Q(z[r, p, t] for r in d.R for p in d.P)
                  + dR[t] >= d.mr[t] * sum(d.gen[r, t] for r in d.R) for t in d.T), name="R7_meta_rec")
    m.addConstrs((Q(v[p, k, t] for p in d.P for k in d.K) + dV[t] >= d.mv[t] * sum(d.gen[r, t] for r in d.R)
                  for t in d.T), name="R8_meta_val")
    m.addConstrs((Q(x[r, c, t] for c in d.C) + Q(z[r, p, t] for p in d.P) + dQ[r, t]
                  >= d.qr[r] * sum(d.gen[rp, t] for rp in d.R) for r in d.R for t in d.T), name="R9_cuota_region")
    for t in d.T[1:]:
        m.addConstrs((y[c, t] >= y[c, anterior[t]] for c in d.C), name=f"R10a_permanencia_{t}")
        m.addConstrs((w[p, k, t] >= w[p, k, anterior[t]] for p in d.P for k in d.K), name=f"R10b_permanencia_{t}")
    # R11: dominios declarados al crear las variables. R12 es nacional y anual.
    m.addConstrs((Q(v[p, k, t] for p in d.P for k in d.KM) >= d.alpha * Q(v[p, k, t] for p in d.P for k in d.K)
                  for t in d.T), name="R12_composicion_material")
    m.update()
    variables = dict(y=y, w=w, x=x, z=z, f=f, v=v, s=s, dR=dR, dV=dV, dQ=dQ)
    return m, variables, costos, anterior


def finito(valor):
    """Los centinelas de infinito de Gurobi se exportan como dato no disponible."""
    if valor is None:
        return None
    valor = float(valor)
    return valor if math.isfinite(valor) and abs(valor) < 1e90 else None


def atributo(modelo, nombre):
    try:
        return finito(modelo.getAttr(nombre))
    except (AttributeError, gp.GurobiError):
        return None


def progreso_callback(historial):
    ultimo = [-5.0]
    def registrar(modelo, donde):
        if donde != gp.GRB.Callback.MIP:
            return
        tiempo = modelo.cbGet(gp.GRB.Callback.RUNTIME)
        if tiempo - ultimo[0] < 5.0:
            return
        ultimo[0] = tiempo
        obj = finito(modelo.cbGet(gp.GRB.Callback.MIP_OBJBST))
        cota = finito(modelo.cbGet(gp.GRB.Callback.MIP_OBJBND))
        gap = abs(obj-cota) / abs(obj) if obj is not None and cota is not None and obj != 0 else None
        historial.append(dict(tiempo_s=tiempo, objetivo_M_CLP=obj, cota_M_CLP=cota,
            gap_pct=100*gap if gap is not None else None,
            nodos=modelo.cbGet(gp.GRB.Callback.MIP_NODCNT)))
    return registrar


def guardar_resultados(m, d, var, costos, anterior, salida, ruta, historial):
    import json
    hay_solucion = m.SolCount > 0
    obj = atributo(m, "ObjVal") if hay_solucion else None
    cota = atributo(m, "ObjBound")
    gap = atributo(m, "MIPGap") if hay_solucion else None
    info = dict(estado_codigo=m.Status, estado=ESTADOS.get(m.Status, str(m.Status)),
        hay_solucion=hay_solucion, optimalidad_certificada_en_tolerancia=m.Status == gp.GRB.OPTIMAL,
        objetivo_M_CLP=obj, objetivo_CLP=obj*1000 if obj is not None else None,
        cota_M_CLP=cota, cota_CLP=cota*1000 if cota is not None else None,
        gap_fraccion=gap, gap_pct=100*gap if gap is not None else None,
        brecha_absoluta_M_CLP=abs(obj-cota) if obj is not None and cota is not None else None,
        tiempo_solver_s=m.Runtime, limite_s=TIME_LIMIT, tolerancia_gap_fraccion=MIP_GAP,
        nodos=atributo(m, "NodeCount"), iteraciones_simplex=atributo(m, "IterCount"),
        soluciones_almacenadas=m.SolCount, variables=m.NumVars, binarias=m.NumBinVars,
        restricciones=m.NumConstrs, no_ceros=atributo(m, "NumNZs"), trabajo=atributo(m, "Work"),
        violacion_restricciones=atributo(m, "ConstrVio") if hay_solucion else None,
        violacion_cotas=atributo(m, "BoundVio") if hay_solucion else None,
        violacion_integralidad=atributo(m, "IntVio") if hay_solucion else None,
        gurobi_version=".".join(map(str, gp.gurobi.version())), python_version=platform.python_version(),
        sistema=platform.platform(), semilla=52, unidad_costos="M$ = 1000 CLP",
        origen_inversiones=d.origen_inversion, archivo_entrada=ruta.name,
        sha256_entrada=hashlib.sha256(ruta.read_bytes()).hexdigest(),
        fecha_ejecucion=datetime.now().astimezone().isoformat())
    historial.append(dict(tiempo_s=m.Runtime, objetivo_M_CLP=obj, cota_M_CLP=cota,
        gap_pct=info["gap_pct"], nodos=info["nodos"]))
    tablas = {"Historial_Gurobi": pd.DataFrame(historial)}
    resumen = [f"Estado: {info['estado']} (código {m.Status}).",
        f"Tiempo de optimización: {m.Runtime:.3f} s; límite: {TIME_LIMIT:.0f} s.",
        f"Variables: {m.NumVars}; restricciones: {m.NumConstrs}; soluciones almacenadas: {m.SolCount}.",
        f"Origen de inversiones: {d.origen_inversion}."]
    if hay_solucion:
        def X(nombre, *indices):
            clave = indices[0] if len(indices) == 1 else tuple(indices)
            return float(var[nombre][clave].X)
        costos_val = {nombre: float(exp.getValue()) for nombre, exp in costos.items()}
        total = sum(costos_val.values())
        filas_costos = [dict(Componente=k, Costo_M_CLP=v, Costo_CLP=1000*v,
            Pct=100*v/total if total != 0 else None) for k, v in costos_val.items()]
        filas_costos.append(dict(Componente="TOTAL", Costo_M_CLP=total, Costo_CLP=1000*total, Pct=100 if total else None))
        tablas["Resumen_Costos_FO"] = pd.DataFrame(filas_costos)
        ds8, cuotas, acopios, plantas, flujos, aperturas, balances = [], [], [], [], [], [], []
        nr = d.regiones.Nombre_Region.to_dict()
        nc = d.centros.Nombre_Centro.to_dict()
        np_ = d.plantas.groupby("ID_Planta").Nombre_Planta.first().to_dict()
        nk = d.tecnologias.Nombre_Tecnologia.to_dict()
        for t in d.T:
            gen = sum(d.gen[r, t] for r in d.R)
            rec = sum(X("x", r, c, t) for r in d.R for c in d.C) + sum(X("z", r, p, t) for r in d.R for p in d.P)
            val = sum(X("v", p, k, t) for p in d.P for k in d.K)
            mat = sum(X("v", p, k, t) for p in d.P for k in d.KM)
            stock_previo = sum(X("s", c, anterior[t]) if anterior[t] is not None else d.s0[c] for c in d.C)
            stock_final = sum(X("s", c, t) for c in d.C)
            residuo = rec + stock_previo - val - stock_final
            balances.append(dict(anio=t, recoleccion_ton=rec, inventario_inicial_ton=stock_previo,
                valorizacion_ton=val, inventario_final_ton=stock_final, residuo_balance_ton=residuo))
            ds8.append(dict(anio=t, generacion_ton=gen, meta_rec_ton=d.mr[t]*gen, recoleccion_ton=rec,
                pct_recoleccion=100*rec/gen, deficit_rec_ton=X("dR", t), meta_val_ton=d.mv[t]*gen,
                valorizacion_ton=val, pct_valorizacion=100*val/gen, deficit_val_ton=X("dV", t),
                valorizacion_material_ton=mat, valorizacion_energetica_ton=val-mat,
                pct_material_sobre_valorizado=100*mat/val if val > TOL_TON else None,
                minimo_material_ton=d.alpha*val, holgura_material_ton=mat-d.alpha*val,
                deficit_cuotas_regionales_ton=sum(X("dQ", r, t) for r in d.R)))
            for r in d.R:
                recogido = sum(X("x", r, c, t) for c in d.C) + sum(X("z", r, p, t) for p in d.P)
                cuotas.append(dict(anio=t, region_id=r, region=nr[r], disponible_ton=d.gen[r, t],
                    cuota_minima_ton=d.qr[r]*gen, recoleccion_ton=recogido, deficit_ton=X("dQ", r, t)))
            for c in d.C:
                activo = X("y", c, t)
                apertura = activo - (X("y", c, anterior[t]) if anterior[t] is not None else 0)
                entrada = sum(X("x", r, c, t) for r in d.R)
                despacho = sum(X("f", c, p, t) for p in d.P)
                acopios.append(dict(anio=t, centro_id=c, centro=nc[c], activo=round(activo),
                    cap_entrada_ton=d.aC[c]*activo, flujo_entrada_ton=entrada,
                    uso_entrada_pct=100*entrada/(d.aC[c]*activo) if d.aC[c]*activo > TOL_TON else None,
                    cap_salida_ton=d.bC[c]*activo, flujo_salida_ton=despacho,
                    uso_salida_pct=100*despacho/(d.bC[c]*activo) if d.bC[c]*activo > TOL_TON else None,
                    inv_max_ton=d.sM[c]*activo, inventario_final_ton=X("s", c, t)))
                aperturas.append(dict(anio=t, tipo="Centro", instalacion=c, tecnologia="",
                    activo=round(activo), apertura=round(apertura), inversion_M_CLP=d.iC[c]*apertura))
            for p in d.P:
                for k in d.K:
                    activo = X("w", p, k, t)
                    apertura = activo - (X("w", p, k, anterior[t]) if anterior[t] is not None else 0)
                    proc = X("v", p, k, t)
                    plantas.append(dict(anio=t, planta_id=p, planta=np_[p], tecnologia_id=k, tecnologia=nk[k],
                        categoria_ds8="Material" if k in d.KM else "Energetica", compatible=d.epk[p, k],
                        activa=round(activo), cap_efectiva_ton=d.aP[p, k]*activo, valorizado_ton=proc,
                        uso_capacidad_pct=100*proc/(d.aP[p, k]*activo) if d.aP[p, k]*activo > TOL_TON else None))
                    aperturas.append(dict(anio=t, tipo="Linea de planta", instalacion=p, tecnologia=k,
                        activo=round(activo), apertura=round(apertura), inversion_M_CLP=d.iP[p, k]*apertura))
            for nombre, orig, dest, etiquetas_o, etiquetas_d, costo, dist in [
                ("x", d.R, d.C, nr, nc, d.cT1, d.distT1),
                ("f", d.C, d.P, nc, np_, d.cT2, d.distT2),
                ("z", d.R, d.P, nr, np_, d.cT3, d.distT3)]:
                for o in orig:
                    for q in dest:
                        flujo = X(nombre, o, q, t)
                        if flujo > TOL_TON:
                            flujos.append(dict(anio=t, tramo={"x":"T1_Region_Acopio", "f":"T2_Acopio_Planta", "z":"T3_Region_Planta_Directo"}[nombre],
                                origen_id=o, origen=etiquetas_o[o], destino_id=q, destino=etiquetas_d[q],
                                distancia_km=dist[o, q], costo_unit_M_CLP=costo[o, q], toneladas=flujo))
        tablas.update({"Cumplimiento_Normativo_DS8":pd.DataFrame(ds8), "Cuotas_Regionales":pd.DataFrame(cuotas),
            "Capacidad_Centros_Acopio":pd.DataFrame(acopios), "Capacidad_Plantas_Tec":pd.DataFrame(plantas),
            "Matriz_Flujos_Transporte":pd.DataFrame(flujos, columns=["anio","tramo","origen_id","origen","destino_id","destino","distancia_km","costo_unit_M_CLP","toneladas"]),
            "Aperturas_Inversion":pd.DataFrame(aperturas), "Balance_Global":pd.DataFrame(balances)})
        # Incluye TODOS los valores, también ceros: respaldo íntegro de las variables.
        tablas["Variables_Modelo"] = pd.DataFrame([
            dict(variable=variable.VarName, valor=variable.X, unidad="adimensional" if nombre in ["y", "w"] else "ton")
            for nombre, grupo in var.items() for variable in grupo.values()])
        restricciones = []
        for restr in m.getConstrs():
            slack = float(restr.Slack)
            vio = abs(slack) if restr.Sense == "=" else max(0.0, -slack) if restr.Sense == "<" else max(0.0, slack)
            restricciones.append(dict(restriccion=restr.ConstrName, sentido=restr.Sense,
                lado_derecho=restr.RHS, holgura=slack, violacion=vio))
        tablas["Verificacion_Restricciones"] = pd.DataFrame(restricciones)
        info["diferencia_reconstruccion_FO_M_CLP"] = abs(total-obj)
        info["max_residuo_balance_global_ton"] = max(abs(b["residuo_balance_ton"]) for b in balances)
        info["deficit_total_rec_ton"] = sum(row["deficit_rec_ton"] for row in ds8)
        info["deficit_total_val_ton"] = sum(row["deficit_val_ton"] for row in ds8)
        info["deficit_total_cuotas_ton"] = sum(row["deficit_ton"] for row in cuotas)
        info["cumple_normativa_sin_deficits"] = (
            max([X("dR", t) for t in d.T] + [X("dV", t) for t in d.T] + [X("dQ", r, t) for r in d.R for t in d.T]) <= TOL_TON
            and min(row["holgura_material_ton"] for row in ds8) >= -TOL_TON)
        resumen += [f"Mejor solución: {obj:,.6f} M$ ({obj*1000:,.2f} CLP).",
            f"Mejor cota: {cota:,.6f} M$." if cota is not None else "Mejor cota: no disponible.",
            f"GAP relativo: {100*gap:.6f} %." if gap is not None else "GAP relativo: no disponible.",
            "Optimalidad certificada dentro de la tolerancia configurada." if m.Status == gp.GRB.OPTIMAL
            else "Se conserva la mejor solución hallada; no se certificó optimalidad dentro de la tolerancia.",
            f"Cumplimiento normativo sin déficits: {info['cumple_normativa_sin_deficits']}.",
            f"Déficits acumulados (ton): recolección={info['deficit_total_rec_ton']:.6f}; "
            f"valorización={info['deficit_total_val_ton']:.6f}; cuotas={info['deficit_total_cuotas_ton']:.6f}.",
            f"Diferencia al reconstruir FO: {abs(total-obj):.8f} M$.",
            f"Mayor residuo de balance global: {info['max_residuo_balance_global_ton']:.8f} ton."]
        m.write(str(salida / "solucion.sol"))
    else:
        resumen.append("No se encontró una solución factible. No se exportan valores de variables ni costos de una solución inexistente.")
        if m.Status == gp.GRB.INF_OR_UNBD:
            resumen.append("INF_OR_UNBD no demuestra infactibilidad: requiere diagnóstico adicional con DualReductions=0.")
    estado = pd.DataFrame([info])
    estado.to_csv(salida / "resumen_solver.csv", index=False, encoding="utf-8-sig")
    (salida / "resumen_solver.json").write_text(json.dumps(info, ensure_ascii=False, indent=2, allow_nan=False), encoding="utf-8")
    csv_nombres = {"Resumen_Costos_FO":"resumen_costos", "Cumplimiento_Normativo_DS8":"cumplimiento_ds8",
        "Cuotas_Regionales":"cuotas_regionales", "Capacidad_Centros_Acopio":"capacidad_acopios",
        "Capacidad_Plantas_Tec":"capacidad_plantas", "Matriz_Flujos_Transporte":"flujos_transporte",
        "Aperturas_Inversion":"aperturas_inversion", "Balance_Global":"balance_global",
        "Variables_Modelo":"variables_modelo", "Verificacion_Restricciones":"verificacion_restricciones",
        "Historial_Gurobi":"historial_gurobi"}
    # Mismo mecanismo de exportación de Excel que el script original del grupo.
    with pd.ExcelWriter(salida / "Resultados_Optimizacion_REP_NFU.xlsx") as writer:
        estado.to_excel(writer, sheet_name="Estado_Gurobi", index=False)
        for hoja, df in tablas.items():
            df.to_excel(writer, sheet_name=hoja, index=False)
            df.to_csv(salida / f"{csv_nombres[hoja]}.csv", index=False, encoding="utf-8-sig")
    texto = "\n".join(resumen) + "\n"
    (salida / "resumen_ejecucion.txt").write_text(texto, encoding="utf-8")
    print("\n" + texto)
    print(f"Archivos guardados en: {salida}")
    return info


def main():
    parser = argparse.ArgumentParser(description="Red REP-NFU, E3, grupo 52")
    parser.add_argument("--datos", help="Ruta opcional del Excel de entrada")
    parser.add_argument("--salida", help="Carpeta raíz opcional de resultados")
    parser.add_argument("--validar-datos", action="store_true", help="Validar parámetros sin optimizar")
    args = parser.parse_args()
    ruta = ruta_datos(args.datos)
    d = cargar_datos(ruta)
    print(f"Excel validado: {ruta}")
    print(f"{len(d.R)} regiones; {len(d.C)} centros; {len(d.P)} plantas; {len(d.K)} tecnologías; años {d.T}.")
    print(f"Inversiones: {d.origen_inversion}. M$ = miles de pesos chilenos.")
    if args.validar_datos:
        return 0
    if gp is None:
        raise ImportError("No está disponible gurobipy. Ejecuta el script con tu intérprete Python 3.11.9 que ya tiene Gurobi 10.0.3.")
    base = Path(__file__).resolve().parent
    raiz = base.parent if base.name in ["Script Python", "Script_Python"] else base
    carpeta = Path(args.salida).expanduser().resolve() if args.salida else raiz / "Analisis de Excels"
    salida = carpeta / datetime.now().astimezone().strftime("ejecucion_%Y%m%d_%H%M%S_%f")
    salida.mkdir(parents=True)
    shutil.copy2(ruta, salida / "entrada_datos.xlsx")
    m = None
    try:
        m, variables, costos, anterior = construir_modelo(d, salida)
        m.write(str(salida / "modelo.lp"))
        m.write(str(salida / "parametros.prm"))
        historial = []
        m.optimize(progreso_callback(historial))
        # No consultar ObjVal, MIPGap o .X sin comprobar primero SolCount.
        info = guardar_resultados(m, d, variables, costos, anterior, salida, ruta, historial)
        if m.Status == gp.GRB.INFEASIBLE:
            # El modelo con déficits admite la solución de flujos nulos; una
            # infactibilidad sugiere errores de datos o de implementación.
            restante = max(0.0, TIME_LIMIT - m.Runtime)
            if restante > 1.0:
                try:
                    m.Params.TimeLimit = restante
                    m.computeIIS()
                    m.write(str(salida / "modelo_infactible.ilp"))
                except gp.GurobiError as error:
                    (salida / "diagnostico_iis.txt").write_text(str(error), encoding="utf-8")
        return 0 if info["hay_solucion"] else 2
    finally:
        if m is not None:
            m.dispose()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (ValueError, FileNotFoundError, ImportError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        sys.exit(1)
