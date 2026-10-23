#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Proyecto REP-NFU, grupo 52 — Sistema de Gestión "Valora Más" (DS N°8 MMA, 2026-2030).
Python 3.11 / gurobipy 10.0.3+.

Modelo actualizado para el nuevo dataset 2026-2030:
- Trazabilidad total de la Región de Origen (Multi-commodity, 16 regiones de Chile).
- Vehículos discretos (Camiones de 10t, 15t y 25t) con variables enteras para viajes.
- Escala Temporal: Anual (removidas las quincenas para optimizar el tiempo de cómputo).
- Cálculo explícito y reporte detallado del Valor Óptimo de la Función Objetivo.
"""



import argparse
import math
import shutil
import sys
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
import pandas as pd
import matplotlib.pyplot as plt  
import seaborn as sns

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
# FUNCIONES DE CARGA Y LECTURA DE EXCEL
# =============================================================================

def ruta_datos(explicita=None):
    """Localiza el archivo Excel de datos (`datos.xlsx`) o valida la ruta explícita provista."""
    base = Path(__file__).resolve().parent

    if explicita:
        ruta = Path(explicita).expanduser().resolve()
        if not ruta.is_file():
            raise FileNotFoundError(f"No existe el archivo Excel de entrada: {ruta}")
        if ruta.suffix.lower() not in [".xlsx", ".xls", ".xlsm"]:
            raise ValueError(f"El archivo debe ser una planilla Excel válida (.xlsx), recibido: '{ruta.name}'")
        return ruta

    candidatos = [
        base / "datos.xlsx",
        base / "dataset" / "datos.xlsx",
        base.parent / "Instancias Excel" / "datos.xlsx",
        base / "Instancias Excel" / "datos.xlsx",
        base.parent / "Instancias_Excel" / "datos.xlsx",
        base / "Instancias_Excel" / "datos.xlsx",
    ]
    for r in candidatos:
        if r.is_file():
            return r

    raise FileNotFoundError(
        "No se encontró 'datos.xlsx' en las ubicaciones esperadas. "
        "Asegúrate de colocar 'datos.xlsx' en el directorio del script o indicar su ubicación con '--datos ruta/al/archivo.xlsx'."
    )


def leer_tabla(libro, hoja, columnas, ids=()):
    df = pd.read_excel(libro, sheet_name=hoja)
    faltantes = [c for c in columnas if c not in df.columns]
    if faltantes:
        raise ValueError(f"En la hoja '{hoja}' faltan las columnas requeridas: {faltantes}")
    if ids:
        df = df.loc[df[ids[0]].notna()].copy()
        for col in ids:
            df[col] = df[col].astype(str).str.strip()
    return df


def numericos(df, columnas):
    for col in columnas:
        df[col] = pd.to_numeric(df[col], errors="coerce")


def cargar_datos(ruta):
    libro = pd.ExcelFile(ruta)

    cols_metas = [
        "Anio_t",
        "Meta_Recoleccion_mR [-]",
        "Meta_Valorizacion_mV [-]",
        "Cuota_Min_Material_alphaM [-]",
        "Penalizacion_muR [M$/ton]",
        "Penalizacion_muV [M$/ton]",
        "Penalizacion_muQ [M$/ton]",
    ]
    metas = leer_tabla(libro, "Metas_DS8", cols_metas)
    numericos(metas, cols_metas)
    metas = metas.set_index("Anio_t").sort_index()
    T = [int(t) for t in metas.index]

    cols_gen = [f"Gen_{t} [ton]" for t in T]
    regiones = leer_tabla(
        libro,
        "Regiones_Generacion",
        ["ID_Region", "Nombre_Region", "Cuota_DS8_qr [-]"] + cols_gen,
        ["ID_Region"],
    )
    regiones = regiones.loc[regiones.ID_Region.str.upper() != "TOTAL"].set_index("ID_Region")
    numericos(regiones, ["Cuota_DS8_qr [-]"] + cols_gen)

    cols_centros_num = [
        "Cap_Recepcion_aC [ton/año]",
        "Cap_Despacho_bC [ton/año]",
        "Limite_Sanitario_sM [ton]",
        "Inv_Inicial_s0 [ton]",
        "Costo_Fijo_Anual_cF [M$/año]",
        "Costo_Almacenaje_h [M$/ton-año]",
        "Inversion_Habilitacion_iC [M$]",
    ]
    centros = leer_tabla(
        libro,
        "Centros_Acopio",
        ["ID_Centro", "Nombre_Centro"] + cols_centros_num,
        ["ID_Centro"],
    ).set_index("ID_Centro")
    numericos(centros, cols_centros_num)

    cols_plantas_num = [
        "Max_Lineas_nP [líneas]",
        "Factibilidad_epk [-]",
        "Cap_Efectiva_aP [ton/año]",
        "Costo_Fijo_cP [M$/año]",
        "Inversion_Habilitacion_iP [M$]",
    ]
    plantas = leer_tabla(
        libro,
        "Plantas_Tecnologias",
        ["ID_Planta", "Nombre_Planta", "ID_Tecnologia"] + cols_plantas_num,
        ["ID_Planta", "ID_Tecnologia"],
    )
    numericos(plantas, cols_plantas_num)

    tecnologias = leer_tabla(
        libro,
        "Tecnologias_Costos",
        ["ID_Tecnologia", "Nombre_Tecnologia", "Subconjunto_DS8", "Costo_Variable_cV [M$/ton]"],
        ["ID_Tecnologia"],
    ).set_index("ID_Tecnologia")
    numericos(tecnologias, ["Costo_Variable_cV [M$/ton]"])

    transporte = leer_tabla(
        libro,
        "Costos_Transporte",
        ["Tipo_Tramo", "Nodo_Origen", "Nodo_Destino", "Distancia_Ruta [km]", "Costo_Unitario [M$/ton]"],
        ["Tipo_Tramo", "Nodo_Origen", "Nodo_Destino"],
    )
    numericos(transporte, ["Distancia_Ruta [km]", "Costo_Unitario [M$/ton]"])

    R, C, K = list(regiones.index), list(centros.index), list(tecnologias.index)
    P = sorted(plantas.ID_Planta.unique())
    categorias = tecnologias.Subconjunto_DS8.astype(str).str.strip()
    KM = list(tecnologias.index[categorias.str.lower() == "material"])

    pares = plantas.set_index(["ID_Planta", "ID_Tecnologia"])
    alpha_dict = metas["Cuota_Min_Material_alphaM [-]"].to_dict()

    d = SimpleNamespace(
        T=T, R=R, C=C, P=P, K=K, KM=KM,
        regiones=regiones, centros=centros, plantas=plantas, tecnologias=tecnologias,
        gen={(r, t): float(regiones.loc[r, f"Gen_{t} [ton]"]) for r in R for t in T},
        qr=regiones["Cuota_DS8_qr [-]"].to_dict(),
        mr=metas["Meta_Recoleccion_mR [-]"].to_dict(),
        mv=metas["Meta_Valorizacion_mV [-]"].to_dict(),
        alpha=float(metas["Cuota_Min_Material_alphaM [-]"].iloc[0]),
        alpha_t=alpha_dict,
        muR=metas["Penalizacion_muR [M$/ton]"].to_dict(),
        muV=metas["Penalizacion_muV [M$/ton]"].to_dict(),
        muQ=metas["Penalizacion_muQ [M$/ton]"].to_dict(),
        aC=centros["Cap_Recepcion_aC [ton/año]"].to_dict(),
        bC=centros["Cap_Despacho_bC [ton/año]"].to_dict(),
        sM=centros["Limite_Sanitario_sM [ton]"].to_dict(),
        s0=centros["Inv_Inicial_s0 [ton]"].to_dict(),
        cF=centros["Costo_Fijo_Anual_cF [M$/año]"].to_dict(),
        h=centros["Costo_Almacenaje_h [M$/ton-año]"].to_dict(),
        iC=centros["Inversion_Habilitacion_iC [M$]"].to_dict(),
        aP=pares["Cap_Efectiva_aP [ton/año]"].to_dict(),
        epk=pares["Factibilidad_epk [-]"].to_dict(),
        cP=pares["Costo_Fijo_cP [M$/año]"].to_dict(),
        iP=pares["Inversion_Habilitacion_iP [M$]"].to_dict(),
        nP=plantas.groupby("ID_Planta")["Max_Lineas_nP [líneas]"].first().to_dict(),
        cV=tecnologias["Costo_Variable_cV [M$/ton]"].to_dict(),
    )

    for tipo, nombre in [("T1_Region_Centro", "T1"), ("T2_Centro_Planta", "T2"), ("T3_Region_Planta", "T3")]:
        tramo = transporte.loc[transporte.Tipo_Tramo == tipo].set_index(["Nodo_Origen", "Nodo_Destino"])
        setattr(d, "c" + nombre, tramo["Costo_Unitario [M$/ton]"].to_dict())
        setattr(d, "dist" + nombre, tramo["Distancia_Ruta [km]"].to_dict())

    d.origen_inversion = "Montos ingresados en Excel (Dataset Real Chile 2024-2026)"
    if "Supuestos_Inversion" in libro.sheet_names:
        supuestos = pd.read_excel(libro, sheet_name="Supuestos_Inversion").set_index("Parametro")
        if "Origen_Inversiones" in supuestos.index:
            d.origen_inversion = str(supuestos.loc["Origen_Inversiones", "Valor"])
    libro.close()
    return d


# =============================================================================
# MODELO MATEMÁTICO GUROBI
# =============================================================================

def construir_modelo_extremo(d, salida, time_limit=TIME_LIMIT, mip_gap=MIP_GAP):
    if gp is None:
        raise ImportError(
            "No se encontró el paquete 'gurobipy'. Instálalo y activa tu licencia Gurobi en tu computador."
        )

    GRB, Q = gp.GRB, gp.quicksum
    m = gp.Model("Optimizacion_REP_NFU_ValoraMas_2026_2030")
    m.Params.LogFile = str(salida / "gurobi.log")
    m.Params.TimeLimit = float(time_limit)
    m.Params.MIPGap = float(mip_gap)
    m.Params.Seed = 52

    VEHICULOS = {
        "Camion_10t": {"cap": 10.0, "costo_f": 1.2},
        "Camion_15t": {"cap": 15.0, "costo_f": 1.1},
        "Camion_25t": {"cap": 25.0, "costo_f": 1.0},
    }
    V = list(VEHICULOS.keys())
    d.V = V

    # Variables de decisión binarias de habilitación / operación anual
    y = m.addVars(d.C, d.T, vtype=GRB.BINARY, name="y")
    w = m.addVars(d.P, d.K, d.T, vtype=GRB.BINARY, name="w")

    # Variables enteras de número de viajes por tipo de camión
    nx = m.addVars(d.R, d.C, V, d.T, vtype=GRB.INTEGER, lb=0, name="nx")
    nz = m.addVars(d.R, d.P, V, d.T, vtype=GRB.INTEGER, lb=0, name="nz")
    nf = m.addVars(d.R, d.C, d.P, V, d.T, vtype=GRB.INTEGER, lb=0, name="nf")

    # Variables continuas de flujo (toneladas), procesamiento, inventario y déficits
    x = m.addVars(d.R, d.C, V, d.T, lb=0, name="x")
    z = m.addVars(d.R, d.P, V, d.T, lb=0, name="z")
    f = m.addVars(d.R, d.C, d.P, V, d.T, lb=0, name="f")
    v = m.addVars(d.R, d.P, d.K, d.T, lb=0, name="v")
    s = m.addVars(d.R, d.C, d.T, lb=0, name="s")

    dR = m.addVars(d.T, lb=0, name="dR")
    dV = m.addVars(d.T, lb=0, name="dV")
    dQ = m.addVars(d.R, d.T, lb=0, name="dQ")

    anterior = {t: d.T[i - 1] if i else None for i, t in enumerate(d.T)}
    dy = {(c, t): y[c, t] - (y[c, anterior[t]] if anterior[t] is not None else 0) for c in d.C for t in d.T}
    dw = {(p, k, t): w[p, k, t] - (w[p, k, anterior[t]] if anterior[t] is not None else 0) for p in d.P for k in d.K for t in d.T}

    # Componentes de la Función Objetivo (en M$ = miles de CLP)
    costos = {
        "IC_C inversion habilitacion centros": Q(d.iC[c] * dy[c, t] for c in d.C for t in d.T),
        "IC_P inversion habilitacion lineas": Q(d.iP[p, k] * dw[p, k, t] for p in d.P for k in d.K for t in d.T),
        "CF_C operacion fija centros": Q(d.cF[c] * y[c, t] for c in d.C for t in d.T),
        "CF_P operacion fija plantas": Q(d.cP[p, k] * w[p, k, t] for p in d.P for k in d.K for t in d.T),
        "CT_1 transporte region-acopio": Q(
            d.cT1[o, c] * VEHICULOS[vh]["costo_f"] * x[o, c, vh, t]
            for o in d.R for c in d.C for vh in V for t in d.T
        ),
        "CT_2 transporte acopio-planta": Q(
            d.cT2[c, p] * VEHICULOS[vh]["costo_f"] * f[o, c, p, vh, t]
            for o in d.R for c in d.C for p in d.P for vh in V for t in d.T
        ),
        "CT_3 transporte region-planta": Q(
            d.cT3[o, p] * VEHICULOS[vh]["costo_f"] * z[o, p, vh, t]
            for o in d.R for p in d.P for vh in V for t in d.T
        ),
        "CP_V procesamiento": Q(
            d.cV[k] * v[o, p, k, t]
            for o in d.R for p in d.P for k in d.K for t in d.T
        ),
        "CA_S almacenaje": Q(
            d.h[c] * s[o, c, t]
            for o in d.R for c in d.C for t in d.T
        ),
        "CP_M penalizaciones deficit": (
            Q(d.muR[t] * dR[t] + d.muV[t] * dV[t] for t in d.T)
            + Q(d.muQ[t] * dQ[o, t] for o in d.R for t in d.T)
        ),
    }

    m.setObjective(Q(costos.values()), GRB.MINIMIZE)

    # Restricciones de capacidad vehicular discreta
    m.addConstrs(
        (x[o, c, vh, t] <= VEHICULOS[vh]["cap"] * nx[o, c, vh, t]
         for o in d.R for c in d.C for vh in V for t in d.T),
        "R_RutaX",
    )
    m.addConstrs(
        (z[o, p, vh, t] <= VEHICULOS[vh]["cap"] * nz[o, p, vh, t]
         for o in d.R for p in d.P for vh in V for t in d.T),
        "R_RutaZ",
    )
    m.addConstrs(
        (f[o, c, p, vh, t] <= VEHICULOS[vh]["cap"] * nf[o, c, p, vh, t]
         for o in d.R for c in d.C for p in d.P for vh in V for t in d.T),
        "R_RutaF",
    )

    # R1: Disponibilidad anual de NFU por región
    m.addConstrs(
        (Q(x[r, c, vh, t] for c in d.C for vh in V)
         + Q(z[r, p, vh, t] for p in d.P for vh in V)
         <= d.gen[r, t]
         for r in d.R for t in d.T),
        "R1",
    )

    # R2a, R2b, R3: Capacidad de recepción, despacho y límite sanitario contra incendios
    m.addConstrs(
        (Q(x[o, c, vh, t] for o in d.R for vh in V) <= d.aC[c] * y[c, t]
         for c in d.C for t in d.T),
        "R2a",
    )
    m.addConstrs(
        (Q(f[o, c, p, vh, t] for o in d.R for p in d.P for vh in V) <= d.bC[c] * y[c, t]
         for c in d.C for t in d.T),
        "R2b",
    )
    m.addConstrs(
        (Q(s[o, c, t] for o in d.R) <= d.sM[c] * y[c, t]
         for c in d.C for t in d.T),
        "R3",
    )

    # R4: Conservación de flujo e inventario anual en Centros de Acopio
    gen_total_t0 = sum(d.gen[r, d.T[0]] for r in d.R)
    for o in d.R:
        prop_o = d.gen[o, d.T[0]] / gen_total_t0 if gen_total_t0 > 0 else (1.0 / len(d.R))
        for c in d.C:
            for t in d.T:
                previo = s[o, c, anterior[t]] if anterior[t] is not None else (d.s0[c] * prop_o)
                m.addConstr(
                    s[o, c, t]
                    == previo
                    + Q(x[o, c, vh, t] for vh in V)
                    - Q(f[o, c, p, vh, t] for p in d.P for vh in V),
                    f"R4_{o}_{c}_{t}",
                )

    # R5: Balance de recepción y procesamiento anual en Plantas por región de origen
    m.addConstrs(
        (Q(v[o, p, k, t] for k in d.K)
         == Q(f[o, c, p, vh, t] for c in d.C for vh in V)
         + Q(z[o, p, vh, t] for vh in V)
         for o in d.R for p in d.P for t in d.T),
        "R5",
    )

    # R6, R6b, R6c: Capacidad técnica efectiva, compatibilidad y máximo de líneas
    m.addConstrs(
        (Q(v[o, p, k, t] for o in d.R) <= d.aP[p, k] * w[p, k, t]
         for p in d.P for k in d.K for t in d.T),
        "R6",
    )
    m.addConstrs(
        (w[p, k, t] <= d.epk[p, k] for p in d.P for k in d.K for t in d.T),
        "R6b",
    )
    m.addConstrs(
        (Q(w[p, k, t] for k in d.K) <= d.nP[p] for p in d.P for t in d.T),
        "R6c",
    )

    # R7, R8, R12 (Metas Nacionales) y R9 (Cuotas Regionales DS8)
    for t in d.T:
        gen_anual_total = sum(d.gen[r, t] for r in d.R)
        alpha_t = d.alpha_t.get(t, d.alpha)
        
        m.addConstr(
            Q(x[o, c, vh, t] for o in d.R for c in d.C for vh in V)
            + Q(z[o, p, vh, t] for o in d.R for p in d.P for vh in V)
            + dR[t]
            >= d.mr[t] * gen_anual_total,
            f"R7_{t}",
        )
        m.addConstr(
            Q(v[o, p, k, t] for o in d.R for p in d.P for k in d.K)
            + dV[t]
            >= d.mv[t] * gen_anual_total,
            f"R8_{t}",
        )
        m.addConstr(
            Q(v[o, p, k, t] for o in d.R for p in d.P for k in d.KM)
            >= alpha_t * Q(v[o, p, k, t] for o in d.R for p in d.P for k in d.K),
            f"R12_{t}",
        )
        for r in d.R:
            m.addConstr(
                Q(x[r, c, vh, t] for c in d.C for vh in V)
                + Q(z[r, p, vh, t] for p in d.P for vh in V)
                + dQ[r, t]
                >= d.qr[r] * gen_anual_total,
                f"R9_{r}_{t}",
            )

    # R10a, R10b: Permanencia de apertura de centros y líneas habilitadas
    for t in d.T[1:]:
        m.addConstrs((y[c, t] >= y[c, anterior[t]] for c in d.C), f"R10a_{t}")
        m.addConstrs((w[p, k, t] >= w[p, k, anterior[t]] for p in d.P for k in d.K), f"R10b_{t}")

    m.update()

    variables = dict(y=y, w=w, x=x, z=z, f=f, v=v, s=s, dR=dR, dV=dV, dQ=dQ, nx=nx, nz=nz, nf=nf)
    return m, variables, costos, anterior


# =============================================================================
# EXTRACCIÓN, CÁLCULO DEL VALOR ÓPTIMO DE LA FUNCIÓN OBJETIVO Y CONSOLIDACIÓN
# =============================================================================

def finito(valor):
    if valor is None:
        return None
    valor = float(valor)
    return valor if math.isfinite(valor) and abs(valor) < 1e90 else None


def atributo(modelo, nombre):
    try:
        return finito(modelo.getAttr(nombre))
    except Exception:
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
        gap = abs(obj - cota) / abs(obj) if obj and cota and obj != 0 else None
        historial.append(
            dict(
                tiempo_s=tiempo,
                objetivo_M_CLP=obj,
                objetivo_MM_CLP=(obj / 1000.0) if obj is not None else None,
                cota_M_CLP=cota,
                gap_pct=100 * gap if gap else None,
                nodos=modelo.cbGet(gp.GRB.Callback.MIP_NODCNT),
            )
        )

    return registrar

def generar_graficos(tablas, salida):
    """Genera y guarda los gráficos requeridos por la rúbrica usando los DataFrames en memoria."""
    # 1. Gráfico de Costos (Torta)
    if "Resumen_Costos_FO" in tablas:
        df_costos = tablas["Resumen_Costos_FO"]
        df_pie = df_costos[(df_costos['Componente'] != 'TOTAL FUNCIÓN OBJETIVO (Z*)') & (df_costos['Costo_MM_CLP'] > 0)]
        
        plt.figure(figsize=(8, 8))
        plt.pie(df_pie['Costo_MM_CLP'], labels=df_pie['Componente'], autopct='%1.1f%%', 
                startangle=140, colors=sns.color_palette("Set2"))
        plt.title('Distribución de Costos del Sistema (2026-2030)')
        plt.savefig(salida / 'grafico_costos.png', bbox_inches='tight')
        plt.close()

    # 2. Gráfico de Metas Normativas (Líneas/Área)
    if "Cumplimiento_Normativo_DS8" in tablas:
        df_norma = tablas["Cumplimiento_Normativo_DS8"]
        plt.figure(figsize=(10, 6))
        plt.plot(df_norma['anio'], df_norma['meta_rec_ton'], 'r--', linewidth=2, label='Meta Recolección (DS8)')
        plt.plot(df_norma['anio'], df_norma['recoleccion_ton'], 'b-', marker='o', label='Recolección Efectiva')
        plt.plot(df_norma['anio'], df_norma['valorizacion_material_ton'], 'g-', marker='s', label='Valorización Material (K2)')
        
        plt.fill_between(df_norma['anio'], df_norma['meta_rec_ton'], df_norma['recoleccion_ton'], color='blue', alpha=0.1)
        
        plt.title('Cumplimiento Normativo: Recolección y Valorización')
        plt.xlabel('Año')
        plt.ylabel('Toneladas')
        plt.xticks(df_norma['anio'])
        plt.legend()
        plt.grid(True, linestyle='--', alpha=0.6)
        plt.savefig(salida / 'grafico_metas.png', bbox_inches='tight')
        plt.close()

def guardar_resultados(m, d, var, costos, anterior, salida, ruta, historial):
    hay_solucion = m.SolCount > 0
    obj_m_clp = atributo(m, "ObjVal") if hay_solucion else None
    obj_mm_clp = (obj_m_clp / 1000.0) if obj_m_clp is not None else None
    obj_clp = (obj_m_clp * 1000.0) if obj_m_clp is not None else None
    cota_m_clp = atributo(m, "ObjBound")
    gap = atributo(m, "MIPGap") if hay_solucion else None

    info = dict(
        estado_codigo=m.Status,
        estado=ESTADOS.get(m.Status, str(m.Status)),
        hay_solucion=hay_solucion,
        optimalidad_certificada_en_tolerancia=(m.Status == gp.GRB.OPTIMAL),
        objetivo_M_CLP=obj_m_clp,
        objetivo_MM_CLP=obj_mm_clp,
        objetivo_CLP=obj_clp,
        cota_M_CLP=cota_m_clp,
        cota_MM_CLP=(cota_m_clp / 1000.0) if cota_m_clp is not None else None,
        gap_pct=100 * gap if gap is not None else None,
        tiempo_solver_s=m.Runtime,
        limite_s=m.Params.TimeLimit,
        nodos=atributo(m, "NodeCount"),
        variables=m.NumVars,
        binarias=m.NumBinVars,
        enteras=m.NumIntVars - m.NumBinVars,
        restricciones=m.NumConstrs,
    )

    if len(historial) == 0 or historial[-1].get("tiempo_s", 0) != m.Runtime:
        historial.append(
            dict(
                tiempo_s=m.Runtime,
                objetivo_M_CLP=obj_m_clp,
                objetivo_MM_CLP=obj_mm_clp,
                cota_M_CLP=cota_m_clp,
                gap_pct=info.get("gap_pct"),
                nodos=info.get("nodos"),
            )
        )

    tablas = {"Historial_Gurobi": pd.DataFrame(historial)}
    resumen = [
        "=" * 78,
        "RESULTADO DE LA OPTIMIZACIÓN GUROBI — SISTEMA VALORA MÁS (LEY REP DS N°8)",
        "=" * 78,
        f"Estado del solver           : {info['estado']} (código {m.Status})",
        f"Tiempo de optimización      : {m.Runtime:.3f} s (límite: {m.Params.TimeLimit:.0f} s)",
        f"Variables totales           : {m.NumVars:,} (Binarias: {m.NumBinVars:,}, Enteras: {m.NumIntVars - m.NumBinVars:,}, Continuas: {m.NumVars - m.NumIntVars:,})",
        f"Restricciones totales       : {m.NumConstrs:,} | Soluciones halladas: {m.SolCount}",
    ]

    if hay_solucion:
        resumen.extend([
            "-" * 78,
            f"VALOR ÓPTIMO DE LA FUNCIÓN OBJETIVO (Z*):",
            f"  • En Miles de Pesos (M$)    : {obj_m_clp:,.2f} M$",
            f"  • En Millones de Pesos (MM$): {obj_mm_clp:,.4f} MM$ (millones de CLP)",
            f"  • En Pesos Chilenos (CLP)   : ${obj_clp:,.0f} CLP",
            f"  • Cota Dual (Lower Bound)   : {cota_m_clp:,.2f} M$ ({(cota_m_clp/1000.0):,.4f} MM$)" if cota_m_clp is not None else "",
            f"  • GAP de Optimalidad        : {info['gap_pct']:.4f}%" if info["gap_pct"] is not None else "",
            "-" * 78,
            "DESGLOSE DE LA FUNCIÓN OBJETIVO POR COMPONENTE DE COSTO:",
        ])

        def X(nombre, *indices):
            clave = indices[0] if len(indices) == 1 else tuple(indices)
            return float(var[nombre][clave].X)

        costos_val = {nombre: float(exp.getValue()) for nombre, exp in costos.items()}
        total = sum(costos_val.values())
        filas_costos = []
        for k, v_m in costos_val.items():
            pct = (100.0 * v_m / total) if total != 0 else 0.0
            filas_costos.append(
                dict(
                    Componente=k,
                    Costo_M_CLP=v_m,
                    Costo_MM_CLP=v_m / 1000.0,
                    Costo_CLP=1000.0 * v_m,
                    Pct=pct,
                )
            )
            resumen.append(
                f"  - {k:<38s}: {v_m:>14,.2f} M$ | {(v_m/1000.0):>10,.2f} MM$ | ${v_m*1000.0:>16,.0f} CLP ({pct:>5.2f}%)"
            )

        filas_costos.append(
            dict(
                Componente="TOTAL FUNCIÓN OBJETIVO (Z*)",
                Costo_M_CLP=total,
                Costo_MM_CLP=total / 1000.0,
                Costo_CLP=1000.0 * total,
                Pct=100.0 if total else None,
            )
        )
        resumen.append("=" * 78)
        tablas["Resumen_Costos_FO"] = pd.DataFrame(filas_costos)

        ds8, cuotas, acopios, plantas, flujos, aperturas, balances = [], [], [], [], [], [], []
        nr = d.regiones.Nombre_Region.to_dict()
        nc = d.centros.Nombre_Centro.to_dict()
        np_ = d.plantas.groupby("ID_Planta").Nombre_Planta.first().to_dict()
        nk = d.tecnologias.Nombre_Tecnologia.to_dict()

        for t in d.T:
            gen = sum(d.gen[r, t] for r in d.R)
            alpha_t = d.alpha_t.get(t, d.alpha)

            rec = (
                sum(X("x", o, c, vh, t) for o in d.R for c in d.C for vh in d.V)
                + sum(X("z", o, p, vh, t) for o in d.R for p in d.P for vh in d.V)
            )
            val = sum(X("v", o, p, k, t) for o in d.R for p in d.P for k in d.K)
            mat = sum(X("v", o, p, k, t) for o in d.R for p in d.P for k in d.KM)

            stock_previo = (
                sum(X("s", o, c, anterior[t]) for o in d.R for c in d.C)
                if anterior[t] is not None
                else sum(d.s0[c] for c in d.C)
            )
            stock_final = sum(X("s", o, c, t) for o in d.R for c in d.C)
            residuo = rec + stock_previo - val - stock_final

            balances.append(
                dict(
                    anio=t,
                    recoleccion_ton=rec,
                    inventario_inicial_ton=stock_previo,
                    valorizacion_ton=val,
                    inventario_final_ton=stock_final,
                    residuo_balance_ton=residuo,
                )
            )

            ds8.append(
                dict(
                    anio=t,
                    generacion_ton=gen,
                    meta_rec_ton=d.mr[t] * gen,
                    recoleccion_ton=rec,
                    pct_recoleccion=100.0 * rec / gen,
                    deficit_rec_ton=X("dR", t),
                    meta_val_ton=d.mv[t] * gen,
                    valorizacion_ton=val,
                    pct_valorizacion=100.0 * val / gen,
                    deficit_val_ton=X("dV", t),
                    valorizacion_material_ton=mat,
                    valorizacion_energetica_ton=val - mat,
                    pct_material_sobre_valorizado=100.0 * mat / val if val > TOL_TON else None,
                    minimo_material_ton=alpha_t * val,
                    holgura_material_ton=mat - alpha_t * val,
                    deficit_cuotas_regionales_ton=sum(X("dQ", r, t) for r in d.R),
                )
            )

            for r in d.R:
                recogido = (
                    sum(X("x", r, c, vh, t) for c in d.C for vh in d.V)
                    + sum(X("z", r, p, vh, t) for p in d.P for vh in d.V)
                )
                cuotas.append(
                    dict(
                        anio=t,
                        region_id=r,
                        region=nr[r],
                        disponible_ton=d.gen[r, t],
                        cuota_minima_ton=d.qr[r] * gen,
                        recoleccion_ton=recogido,
                        deficit_ton=X("dQ", r, t),
                    )
                )

            for c in d.C:
                activo = X("y", c, t)
                apertura = activo - (X("y", c, anterior[t]) if anterior[t] is not None else 0)
                entrada = sum(X("x", o, c, vh, t) for o in d.R for vh in d.V)
                despacho = sum(X("f", o, c, p, vh, t) for o in d.R for p in d.P for vh in d.V)
                acopios.append(
                    dict(
                        anio=t,
                        centro_id=c,
                        centro=nc[c],
                        activo=round(activo),
                        cap_entrada_ton=d.aC[c] * activo,
                        flujo_entrada_ton=entrada,
                        uso_entrada_pct=100.0 * entrada / (d.aC[c] * activo) if d.aC[c] * activo > TOL_TON else None,
                        cap_salida_ton=d.bC[c] * activo,
                        flujo_salida_ton=despacho,
                        uso_salida_pct=100.0 * despacho / (d.bC[c] * activo) if d.bC[c] * activo > TOL_TON else None,
                        inv_max_ton=d.sM[c] * activo,
                        inventario_final_ton=sum(X("s", o, c, t) for o in d.R),
                    )
                )
                aperturas.append(
                    dict(
                        anio=t,
                        tipo="Centro",
                        instalacion=c,
                        tecnologia="",
                        activo=round(activo),
                        apertura=round(apertura),
                        inversion_M_CLP=d.iC[c] * apertura,
                    )
                )

            for p in d.P:
                for k in d.K:
                    activo = X("w", p, k, t)
                    apertura = activo - (X("w", p, k, anterior[t]) if anterior[t] is not None else 0)
                    proc = sum(X("v", o, p, k, t) for o in d.R)
                    plantas.append(
                        dict(
                            anio=t,
                            planta_id=p,
                            planta=np_[p],
                            tecnologia_id=k,
                            tecnologia=nk[k],
                            categoria_ds8="Material" if k in d.KM else "Energetica",
                            compatible=d.epk[p, k],
                            activa=round(activo),
                            cap_efectiva_ton=d.aP[p, k] * activo,
                            valorizado_ton=proc,
                            uso_capacidad_pct=100.0 * proc / (d.aP[p, k] * activo) if d.aP[p, k] * activo > TOL_TON else None,
                        )
                    )
                    aperturas.append(
                        dict(
                            anio=t,
                            tipo="Linea de planta",
                            instalacion=p,
                            tecnologia=k,
                            activo=round(activo),
                            apertura=round(apertura),
                            inversion_M_CLP=d.iP[p, k] * apertura,
                        )
                    )

            for r in d.R:
                for c in d.C:
                    flujo = sum(X("x", r, c, vh, t) for vh in d.V)
                    if flujo > TOL_TON:
                        flujos.append(
                            dict(
                                anio=t,
                                tramo="T1_Region_Acopio",
                                origen_id=r,
                                origen=nr[r],
                                destino_id=c,
                                destino=nc[c],
                                distancia_km=d.distT1[r, c],
                                costo_unit_M_CLP=d.cT1[r, c],
                                toneladas=flujo,
                            )
                        )
            for c in d.C:
                for p in d.P:
                    flujo = sum(X("f", o, c, p, vh, t) for o in d.R for vh in d.V)
                    if flujo > TOL_TON:
                        flujos.append(
                            dict(
                                anio=t,
                                tramo="T2_Acopio_Planta",
                                origen_id=c,
                                origen=nc[c],
                                destino_id=p,
                                destino=np_[p],
                                distancia_km=d.distT2[c, p],
                                costo_unit_M_CLP=d.cT2[c, p],
                                toneladas=flujo,
                            )
                        )
            for r in d.R:
                for p in d.P:
                    flujo = sum(X("z", r, p, vh, t) for vh in d.V)
                    if flujo > TOL_TON:
                        flujos.append(
                            dict(
                                anio=t,
                                tramo="T3_Region_Planta_Directo",
                                origen_id=r,
                                origen=nr[r],
                                destino_id=p,
                                destino=np_[p],
                                distancia_km=d.distT3[r, p],
                                costo_unit_M_CLP=d.cT3[r, p],
                                toneladas=flujo,
                            )
                        )

        tablas.update({
            "Cumplimiento_Normativo_DS8": pd.DataFrame(ds8),
            "Cuotas_Regionales": pd.DataFrame(cuotas),
            "Capacidad_Centros_Acopio": pd.DataFrame(acopios),
            "Capacidad_Plantas_Tec": pd.DataFrame(plantas),
            "Matriz_Flujos_Transporte": pd.DataFrame(
                flujos,
                columns=["anio", "tramo", "origen_id", "origen", "destino_id", "destino", "distancia_km", "costo_unit_M_CLP", "toneladas"],
            ),
            "Aperturas_Inversion": pd.DataFrame(aperturas),
            "Balance_Global": pd.DataFrame(balances),
        })

        m.write(str(salida / "solucion.sol"))

    estado = pd.DataFrame([info])
    estado.to_csv(salida / "resumen_solver.csv", index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(salida / "Resultados_Optimizacion_REP_NFU_Extremo.xlsx") as writer:
        estado.to_excel(writer, sheet_name="Estado_Gurobi", index=False)
        for hoja, df in tablas.items():
            df.to_excel(writer, sheet_name=hoja, index=False)

    try:
        generar_graficos(tablas, salida)
    except Exception as e:
        print(f"Error al generar gráficos: {e}")

    texto = "\n".join(line for line in resumen if line != "") + "\n"
    (salida / "resumen_ejecucion.txt").write_text(texto, encoding="utf-8")

    print("\n" + texto)
    print(f"Archivos guardados en: {salida}")
    return info


# =============================================================================
# MAIN EXECUTOR
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Red REP-NFU ('Valora Más' 2026-2030) - Modelo Gurobi de Complejidad Extrema (Anual)"
    )
    parser.add_argument(
        "-d", "--datos",
        type=str,
        default=None,
        help="Ruta explícita al archivo Excel de entrada (por defecto busca 'datos.xlsx').",
    )
    parser.add_argument(
        "--time-limit",
        type=float,
        default=TIME_LIMIT,
        help=f"Límite de tiempo de Gurobi en segundos (por defecto: {TIME_LIMIT}).",
    )
    parser.add_argument(
        "--mip-gap",
        type=float,
        default=MIP_GAP,
        help=f"Tolerancia de optimalidad relativa MIPGap (por defecto: {MIP_GAP}).",
    )
    args = parser.parse_args()

    ruta = ruta_datos(explicita=args.datos)
    d = cargar_datos(ruta)
    print(
        f"Dataset cargado desde '{ruta.name}': "
        f"{len(d.R)} regiones; {len(d.C)} centros; {len(d.P)} plantas; {len(d.K)} tecnologías; años {d.T}."
    )

    base = Path(__file__).resolve().parent
    raiz = base.parent if base.name in ["Script Python", "Script_Python"] else base
    carpeta = raiz / "Analisis_de_Excels"
    salida = carpeta / datetime.now().astimezone().strftime("ejecucion_%Y%m%d_%H%M%S_%f")
    salida.mkdir(parents=True, exist_ok=True)
    shutil.copy2(ruta, salida / "entrada_datos.xlsx")

    m = None
    try:
        m, variables, costos, anterior = construir_modelo_extremo(
            d, salida, time_limit=args.time_limit, mip_gap=args.mip_gap
        )
        print(f"\n--- ESTADÍSTICAS DEL MODELO DE COMPLEJIDAD EXTREMA ---")
        print(f"Variables totales    : {m.NumVars:,}")
        print(f"  Binarias           : {m.NumBinVars:,}")
        print(f"  Enteras            : {m.NumIntVars - m.NumBinVars:,}")
        print(f"  Continuas          : {m.NumVars - m.NumIntVars:,}")
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