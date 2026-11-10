#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Proyecto REP-NFU, grupo 52 — Sistema de Gestión "Valora Más" (DS N°8 MMA, 2026-2030).
Python 3.11+ / gurobipy 10.0.3+.

Librerías externas usadas: gurobipy, pandas, matplotlib (todas permitidas por el enunciado).
La lectura/escritura de planillas .xlsx se hace con `xlsx_io.py` (solo librería estándar),
por lo que NO se requiere openpyxl.

Cambios respecto de la Entrega 3 (retroalimentación de la ayudante):
- El transporte se cobra por VIAJE (camión completo), no por tonelada: los viajes enteros
  n^X, n^Z, n^F ahora tienen costo y la elección del tipo de camión es una decisión real.
- Tope de flota en regiones extremas (camión-días disponibles por región y tipo de camión).
- Capacidad por viaje, factor tarifario, parámetros operativos y tope de flota se leen
  desde el Excel de entrada (hojas Flota_Camiones, Parametros_Flota y Tope_Flota).
- Los viajes centro-planta ya no se separan por región de origen (carga consolidada).
"""

import argparse
import math
import shutil
import sys
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

import pandas as pd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

from xlsx_io import LibroXlsx, escribir_xlsx

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

CARPETA_DATOS = "Instancias Excel"
CARPETA_SALIDA = "Analisis de Excels"


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
        if ruta.suffix.lower() not in [".xlsx", ".xlsm"]:
            raise ValueError(f"El archivo debe ser una planilla Excel .xlsx, recibido: '{ruta.name}'")
        return ruta

    candidatos = [
        base.parent / CARPETA_DATOS / "datos.xlsx",
        base / CARPETA_DATOS / "datos.xlsx",
        base / "datos.xlsx",
    ]
    for r in candidatos:
        if r.is_file():
            return r

    raise FileNotFoundError(
        f"No se encontró 'datos.xlsx' en '{CARPETA_DATOS}/'. "
        "Indica su ubicación con '--datos ruta/al/archivo.xlsx'."
    )


def leer_tabla(libro, hoja, columnas, ids=()):
    df = libro.hoja(hoja)
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
    libro = LibroXlsx(ruta)

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
    metas = metas.loc[metas.Anio_t.notna()]
    metas["Anio_t"] = metas.Anio_t.astype(int)
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
        ["ID_Centro", "Nombre_Centro", "ID_Region_Ubicacion"] + cols_centros_num,
        ["ID_Centro", "ID_Region_Ubicacion"],
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

    # --- Flota de camiones (antes escrita a mano en el código; ahora desde Excel) ---
    flota = leer_tabla(
        libro,
        "Flota_Camiones",
        ["ID_Camion", "Cap_Viaje_qV [ton/viaje]", "Factor_Tarifario_fV [-]"],
        ["ID_Camion"],
    ).set_index("ID_Camion")
    numericos(flota, ["Cap_Viaje_qV [ton/viaje]", "Factor_Tarifario_fV [-]"])

    pflota = leer_tabla(libro, "Parametros_Flota", ["Parametro", "Valor"], ["Parametro"]).set_index("Parametro")
    numericos(pflota, ["Valor"])
    requeridos = ["Velocidad_Media_v", "Horas_Conduccion_Diarias_H", "Tiempo_Carga_Descarga_tCD", "Dias_Operativos_D"]
    faltan = [p for p in requeridos if p not in pflota.index]
    if faltan:
        raise ValueError(f"En la hoja 'Parametros_Flota' faltan los parámetros: {faltan}")

    V = list(flota.index)
    tope_df = leer_tabla(
        libro,
        "Tope_Flota",
        ["ID_Region"] + [f"Camiones_{u} [un]" for u in V],
        ["ID_Region"],
    ).set_index("ID_Region")
    numericos(tope_df, [f"Camiones_{u} [un]" for u in V])

    R, C, K = list(regiones.index), list(centros.index), list(tecnologias.index)
    P = sorted(plantas.ID_Planta.unique())
    categorias = tecnologias.Subconjunto_DS8.astype(str).str.strip()
    KM = list(tecnologias.index[categorias.str.lower() == "material"])

    pares = plantas.set_index(["ID_Planta", "ID_Tecnologia"])
    alpha_dict = metas["Cuota_Min_Material_alphaM [-]"].to_dict()

    d = SimpleNamespace(
        T=T, R=R, C=C, P=P, K=K, KM=KM, V=V,
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
        region_centro=centros["ID_Region_Ubicacion"].to_dict(),
        aP=pares["Cap_Efectiva_aP [ton/año]"].to_dict(),
        epk=pares["Factibilidad_epk [-]"].to_dict(),
        cP=pares["Costo_Fijo_cP [M$/año]"].to_dict(),
        iP=pares["Inversion_Habilitacion_iP [M$]"].to_dict(),
        nP=plantas.groupby("ID_Planta")["Max_Lineas_nP [líneas]"].first().to_dict(),
        cV=tecnologias["Costo_Variable_cV [M$/ton]"].to_dict(),
        qV=flota["Cap_Viaje_qV [ton/viaje]"].to_dict(),
        fV=flota["Factor_Tarifario_fV [-]"].to_dict(),
        vel=float(pflota.loc["Velocidad_Media_v", "Valor"]),
        horas=float(pflota.loc["Horas_Conduccion_Diarias_H", "Valor"]),
        tCD=float(pflota.loc["Tiempo_Carga_Descarga_tCD", "Valor"]),
        dias=float(pflota.loc["Dias_Operativos_D", "Valor"]),
        tope={
            (r, u): float(tope_df.loc[r, f"Camiones_{u} [un]"])
            for r in tope_df.index if r in R
            for u in V
            if pd.notna(tope_df.loc[r, f"Camiones_{u} [un]"])
        },
    )
    d.RE = sorted({r for r, _ in d.tope})  # regiones con tope de flota

    for tipo, nombre in [("T1_Region_Centro", "T1"), ("T2_Centro_Planta", "T2"), ("T3_Region_Planta", "T3")]:
        tramo = transporte.loc[transporte.Tipo_Tramo == tipo].set_index(["Nodo_Origen", "Nodo_Destino"])
        costos = tramo["Costo_Unitario [M$/ton]"].to_dict()
        dist = tramo["Distancia_Ruta [km]"].to_dict()
        setattr(d, "c" + nombre, costos)
        setattr(d, "dist" + nombre, dist)
        # Duración de un viaje redondo [días-camión]: ida y vuelta + carga/descarga
        setattr(d, "tau" + nombre, {k: 2.0 * km / (d.vel * d.horas) + d.tCD for k, km in dist.items()})

    d.origen_inversion = "Montos ingresados en Excel (Dataset Real Chile 2024-2026)"
    if "Supuestos_Inversion" in libro.sheet_names:
        supuestos = libro.hoja("Supuestos_Inversion").set_index("Parametro")
        if "Origen_Inversiones" in supuestos.index:
            d.origen_inversion = str(supuestos.loc["Origen_Inversiones", "Valor"])
    return d


# =============================================================================
# MODELO MATEMÁTICO GUROBI
# =============================================================================

def construir_modelo(d, salida, time_limit=TIME_LIMIT, mip_gap=MIP_GAP):
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

    V, qV, fV = d.V, d.qV, d.fV

    # Variables binarias de habilitación / operación anual
    y = m.addVars(d.C, d.T, vtype=GRB.BINARY, name="y")
    w = m.addVars(d.P, d.K, d.T, vtype=GRB.BINARY, name="w")

    # Variables enteras de número de viajes por tipo de camión (cotas superiores válidas)
    nx = m.addVars(
        d.R, d.C, V, d.T, vtype=GRB.INTEGER, lb=0,
        ub={(r, c, u, t): math.ceil(d.gen[r, t] / qV[u]) for r in d.R for c in d.C for u in V for t in d.T},
        name="nx",
    )
    nz = m.addVars(
        d.R, d.P, V, d.T, vtype=GRB.INTEGER, lb=0,
        ub={(r, p, u, t): math.ceil(d.gen[r, t] / qV[u]) for r in d.R for p in d.P for u in V for t in d.T},
        name="nz",
    )
    nf = m.addVars(
        d.C, d.P, V, d.T, vtype=GRB.INTEGER, lb=0,
        ub={(c, p, u, t): math.ceil(d.bC[c] / qV[u]) for c in d.C for p in d.P for u in V for t in d.T},
        name="nf",
    )

    # Variables continuas de flujo (toneladas), procesamiento, inventario y déficits
    x = m.addVars(d.R, d.C, d.T, lb=0, name="x")
    z = m.addVars(d.R, d.P, d.T, lb=0, name="z")
    f = m.addVars(d.R, d.C, d.P, d.T, lb=0, name="f")
    v = m.addVars(d.R, d.P, d.K, d.T, lb=0, name="v")
    s = m.addVars(d.R, d.C, d.T, lb=0, name="s")

    dR = m.addVars(d.T, lb=0, name="dR")
    dV = m.addVars(d.T, lb=0, name="dV")
    dQ = m.addVars(d.R, d.T, lb=0, name="dQ")

    anterior = {t: d.T[i - 1] if i else None for i, t in enumerate(d.T)}
    dy = {(c, t): y[c, t] - (y[c, anterior[t]] if anterior[t] is not None else 0) for c in d.C for t in d.T}
    dw = {(p, k, t): w[p, k, t] - (w[p, k, anterior[t]] if anterior[t] is not None else 0) for p in d.P for k in d.K for t in d.T}

    # Costo de un viaje del camión u en un tramo = q_u * f_u * c_tramo [M$/viaje]
    # (se paga el camión completo, vaya lleno o no)
    costos = {
        "IC_C inversion habilitacion centros": Q(d.iC[c] * dy[c, t] for c in d.C for t in d.T),
        "IC_P inversion habilitacion lineas": Q(d.iP[p, k] * dw[p, k, t] for p in d.P for k in d.K for t in d.T),
        "CF_C operacion fija centros": Q(d.cF[c] * y[c, t] for c in d.C for t in d.T),
        "CF_P operacion fija plantas": Q(d.cP[p, k] * w[p, k, t] for p in d.P for k in d.K for t in d.T),
        "CT_1 transporte region-acopio": Q(
            qV[u] * fV[u] * d.cT1[r, c] * nx[r, c, u, t]
            for r in d.R for c in d.C for u in V for t in d.T
        ),
        "CT_2 transporte acopio-planta": Q(
            qV[u] * fV[u] * d.cT2[c, p] * nf[c, p, u, t]
            for c in d.C for p in d.P for u in V for t in d.T
        ),
        "CT_3 transporte region-planta": Q(
            qV[u] * fV[u] * d.cT3[r, p] * nz[r, p, u, t]
            for r in d.R for p in d.P for u in V for t in d.T
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

    # R0: Capacidad vehicular discreta — lo transportado cabe en los viajes contratados
    m.addConstrs(
        (x[r, c, t] <= Q(qV[u] * nx[r, c, u, t] for u in V) for r in d.R for c in d.C for t in d.T),
        "R0_X",
    )
    m.addConstrs(
        (z[r, p, t] <= Q(qV[u] * nz[r, p, u, t] for u in V) for r in d.R for p in d.P for t in d.T),
        "R0_Z",
    )
    m.addConstrs(
        (Q(f[o, c, p, t] for o in d.R) <= Q(qV[u] * nf[c, p, u, t] for u in V)
         for c in d.C for p in d.P for t in d.T),
        "R0_F",
    )

    # R13: Tope de flota en regiones extremas (camión-días disponibles por tipo de camión)
    for r in d.RE:
        centros_r = [c for c in d.C if d.region_centro[c] == r]
        for u in V:
            if (r, u) not in d.tope:
                continue
            for t in d.T:
                m.addConstr(
                    Q(d.tauT1[r, c] * nx[r, c, u, t] for c in d.C)
                    + Q(d.tauT3[r, p] * nz[r, p, u, t] for p in d.P)
                    + Q(d.tauT2[c, p] * nf[c, p, u, t] for c in centros_r for p in d.P)
                    <= d.tope[r, u] * d.dias,
                    f"R13_{r}_{u}_{t}",
                )

    # R1: Disponibilidad anual de NFU por región
    m.addConstrs(
        (Q(x[r, c, t] for c in d.C) + Q(z[r, p, t] for p in d.P) <= d.gen[r, t]
         for r in d.R for t in d.T),
        "R1",
    )

    # R2a, R2b, R3: Capacidad de recepción, despacho y límite sanitario contra incendios
    m.addConstrs(
        (Q(x[o, c, t] for o in d.R) <= d.aC[c] * y[c, t] for c in d.C for t in d.T),
        "R2a",
    )
    m.addConstrs(
        (Q(f[o, c, p, t] for o in d.R for p in d.P) <= d.bC[c] * y[c, t] for c in d.C for t in d.T),
        "R2b",
    )
    m.addConstrs(
        (Q(s[o, c, t] for o in d.R) <= d.sM[c] * y[c, t] for c in d.C for t in d.T),
        "R3",
    )

    # R4: Conservación de flujo e inventario anual en Centros de Acopio (trazable por origen)
    gen_total_t0 = sum(d.gen[r, d.T[0]] for r in d.R)
    for o in d.R:
        prop_o = d.gen[o, d.T[0]] / gen_total_t0 if gen_total_t0 > 0 else (1.0 / len(d.R))
        for c in d.C:
            for t in d.T:
                previo = s[o, c, anterior[t]] if anterior[t] is not None else (d.s0[c] * prop_o)
                m.addConstr(
                    s[o, c, t] == previo + x[o, c, t] - Q(f[o, c, p, t] for p in d.P),
                    f"R4_{o}_{c}_{t}",
                )

    # R5: Balance de recepción y procesamiento anual en Plantas por región de origen
    m.addConstrs(
        (Q(v[o, p, k, t] for k in d.K) == Q(f[o, c, p, t] for c in d.C) + z[o, p, t]
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

    # R14a, R14b: Desigualdades válidas (no eliminan soluciones enteras; ajustan la relajación lineal).
    # Una línea solo procesa NFU de un origen si está activa, y a lo más lo que ese origen genera.
    m.addConstrs(
        (v[o, p, k, t] <= min(d.gen[o, t], d.aP[p, k]) * w[p, k, t]
         for o in d.R for p in d.P for k in d.K for t in d.T),
        "R14a",
    )
    # Solo se despacha directo a una planta que tenga al menos una línea activa.
    m.addConstrs(
        (z[r, p, t] <= d.gen[r, t] * Q(w[p, k, t] for k in d.K)
         for r in d.R for p in d.P for t in d.T),
        "R14b",
    )

    # R7, R8, R12 (Metas Nacionales) y R9 (Cuotas Regionales DS8)
    for t in d.T:
        gen_anual_total = sum(d.gen[r, t] for r in d.R)
        alpha_t = d.alpha_t.get(t, d.alpha)

        m.addConstr(
            Q(x[o, c, t] for o in d.R for c in d.C) + Q(z[o, p, t] for o in d.R for p in d.P) + dR[t]
            >= d.mr[t] * gen_anual_total,
            f"R7_{t}",
        )
        m.addConstr(
            Q(v[o, p, k, t] for o in d.R for p in d.P for k in d.K) + dV[t] >= d.mv[t] * gen_anual_total,
            f"R8_{t}",
        )
        m.addConstr(
            Q(v[o, p, k, t] for o in d.R for p in d.P for k in d.KM)
            >= alpha_t * Q(v[o, p, k, t] for o in d.R for p in d.P for k in d.K),
            f"R12_{t}",
        )
        for r in d.R:
            m.addConstr(
                Q(x[r, c, t] for c in d.C) + Q(z[r, p, t] for p in d.P) + dQ[r, t]
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
    """Genera y guarda los gráficos del informe a partir de los DataFrames en memoria."""
    # 1. Gráfico de Costos (Torta)
    if "Resumen_Costos_FO" in tablas:
        df_costos = tablas["Resumen_Costos_FO"]
        df_pie = df_costos[(df_costos["Componente"] != "TOTAL FUNCIÓN OBJETIVO (Z*)") & (df_costos["Costo_MM_CLP"] > 0)]

        plt.figure(figsize=(8, 8))
        colores = plt.cm.Set2(range(len(df_pie)))
        plt.pie(df_pie["Costo_MM_CLP"], labels=df_pie["Componente"], autopct="%1.1f%%",
                startangle=140, colors=colores)
        plt.title("Distribución de Costos del Sistema (2026-2030)")
        plt.savefig(salida / "grafico_costos.png", bbox_inches="tight")
        plt.close()

    # 2. Gráfico de Metas Normativas
    if "Cumplimiento_Normativo_DS8" in tablas:
        df = tablas["Cumplimiento_Normativo_DS8"]
        plt.figure(figsize=(10, 6))
        plt.plot(df["anio"], df["recoleccion_ton"], "b-", marker="o", linewidth=2, label="Recolección efectiva", zorder=2)
        plt.plot(df["anio"], df["valorizacion_ton"], color="tab:orange", marker="^", linestyle="-", label="Valorización total", zorder=2)
        plt.plot(df["anio"], df["valorizacion_material_ton"], "g-", marker="s", label="Valorización material (K1+K2)", zorder=2)
        # Metas dibujadas encima (zorder alto) para que no queden ocultas cuando coinciden
        plt.plot(df["anio"], df["meta_rec_ton"], "r--", linewidth=2.5, label="Meta recolección (DS8)", zorder=3)
        plt.plot(df["anio"], df["meta_val_ton"], "k:", linewidth=2.5, label="Meta valorización (DS8)", zorder=3)
        plt.plot(df["anio"], df["minimo_material_ton"], color="tab:green", linestyle=":", linewidth=1.8,
                 label="Mínimo material (α·valorizado)", zorder=3)
        plt.title("Cumplimiento Normativo: Recolección y Valorización")
        plt.xlabel("Año")
        plt.ylabel("Toneladas")
        plt.xticks(df["anio"])
        plt.legend()
        plt.grid(True, linestyle="--", alpha=0.6)
        plt.savefig(salida / "grafico_metas.png", bbox_inches="tight")
        plt.close()


def tablas_informe(tablas, d):
    """Construye las tablas procesadas que se reportan en el informe (misma fuente que los resultados)."""
    out = {}
    costos = tablas["Resumen_Costos_FO"].set_index("Componente")
    filas = [
        ("Procesamiento variable en plantas (CP_V)", ["CP_V procesamiento"]),
        ("Operación fija anual de plantas (CF_P)", ["CF_P operacion fija plantas"]),
        ("Transporte directo región-planta (CT_3)", ["CT_3 transporte region-planta"]),
        ("Penalizaciones por déficit normativo (CP_M)", ["CP_M penalizaciones deficit"]),
        ("Inversión en habilitación de líneas (IC_P)", ["IC_P inversion habilitacion lineas"]),
        ("Habilitación, operación fija, flete y acopio en centros",
         ["IC_C inversion habilitacion centros", "CF_C operacion fija centros",
          "CT_1 transporte region-acopio", "CT_2 transporte acopio-planta", "CA_S almacenaje"]),
        ("Total Función Objetivo (Z*)", ["TOTAL FUNCIÓN OBJETIVO (Z*)"]),
    ]
    total = float(costos.loc["TOTAL FUNCIÓN OBJETIVO (Z*)", "Costo_M_CLP"])
    t1 = []
    for nombre, comps in filas:
        val = float(sum(costos.loc[c, "Costo_M_CLP"] for c in comps))
        t1.append({"Componente de costo": nombre, "Costo [M$]": val, "Costo [MM$]": val / 1000.0,
                   "Participación [%]": 100.0 * val / total if total else None})
    out["Tabla 1_Costos_FO"] = pd.DataFrame(t1)

    ds8 = tablas["Cumplimiento_Normativo_DS8"]
    viajes = tablas["Viajes_Flota_Anual"].set_index("anio")
    t2 = pd.DataFrame({
        "Año": ds8.anio,
        "Gen. [ton]": ds8.generacion_ton,
        "Meta Rec. [%]": [100 * d.mr[t] for t in ds8.anio],
        "Rec. [ton]": ds8.recoleccion_ton,
        "Val. [ton]": ds8.valorizacion_ton,
        "% Rec.": ds8.pct_recoleccion,
        "Val. Mat. [ton]": ds8.valorizacion_material_ton,
        "% Mat.": ds8.pct_material_sobre_valorizado,
        "Déficit dQ [ton]": ds8.deficit_cuotas_regionales_ton,
    })
    for u in d.V:
        t2[f"Viajes {u}"] = [viajes.loc[t, f"viajes_{u}"] for t in ds8.anio]
    t2["Viajes total"] = [viajes.loc[t, "viajes_total"] for t in ds8.anio]
    t2["Ocupación media [%]"] = [viajes.loc[t, "ocupacion_media_pct"] for t in ds8.anio]
    out["Tabla 2_Balance_Anual"] = t2

    pl = tablas["Capacidad_Plantas_Tec"]
    activas = pl.loc[pl.activa == 1, ["planta", "tecnologia_id", "categoria_ds8"]].drop_duplicates()
    t3 = []
    for _, fila in activas.iterrows():
        sub = pl.loc[(pl.planta == fila.planta) & (pl.tecnologia_id == fila.tecnologia_id)].set_index("anio")
        reg = {"Planta": fila.planta, "Tecnología": fila.tecnologia_id, "Categoría": fila.categoria_ds8,
               "Cap. [ton]": float(sub.cap_efectiva_ton.max())}
        for t in d.T:
            reg[str(t)] = float(sub.loc[t, "valorizado_ton"]) if sub.loc[t, "activa"] == 1 else None
        t3.append(reg)
    out["Tabla 3_Plantas_Habilitacion"] = pd.DataFrame(t3)
    return out


def guardar_resultados(m, d, var, costos, anterior, salida, historial):
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
            "VALOR ÓPTIMO DE LA FUNCIÓN OBJETIVO (Z*):",
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

        def N(nombre, *indices):
            return int(round(X(nombre, *indices)))

        costos_val = {nombre: float(exp.getValue()) for nombre, exp in costos.items()}
        total = sum(costos_val.values())
        filas_costos = []
        for k, v_m in costos_val.items():
            pct = (100.0 * v_m / total) if total != 0 else 0.0
            filas_costos.append(dict(Componente=k, Costo_M_CLP=v_m, Costo_MM_CLP=v_m / 1000.0,
                                     Costo_CLP=1000.0 * v_m, Pct=pct))
            resumen.append(
                f"  - {k:<38s}: {v_m:>14,.2f} M$ | {(v_m/1000.0):>10,.2f} MM$ | ${v_m*1000.0:>16,.0f} CLP ({pct:>5.2f}%)"
            )
        filas_costos.append(dict(Componente="TOTAL FUNCIÓN OBJETIVO (Z*)", Costo_M_CLP=total,
                                 Costo_MM_CLP=total / 1000.0, Costo_CLP=1000.0 * total,
                                 Pct=100.0 if total else None))
        resumen.append("=" * 78)
        tablas["Resumen_Costos_FO"] = pd.DataFrame(filas_costos)

        ds8, cuotas, acopios, plantas, flujos, aperturas, balances = [], [], [], [], [], [], []
        viajes_anual, uso_flota = [], []
        nr = d.regiones.Nombre_Region.to_dict()
        nc = d.centros.Nombre_Centro.to_dict()
        np_ = d.plantas.groupby("ID_Planta").Nombre_Planta.first().to_dict()
        nk = d.tecnologias.Nombre_Tecnologia.to_dict()

        def registro_flujo(t, tramo, oid, onom, did, dnom, dist, cunit, ton, viajes_u, tau):
            cap = sum(d.qV[u] * viajes_u[u] for u in d.V)
            reg = dict(anio=t, tramo=tramo, origen_id=oid, origen=onom, destino_id=did, destino=dnom,
                       distancia_km=dist, costo_unit_M_CLP_ton=cunit, dias_por_viaje=tau, toneladas=ton)
            for u in d.V:
                reg[f"viajes_{u}"] = viajes_u[u]
            reg["viajes_total"] = sum(viajes_u.values())
            reg["capacidad_contratada_ton"] = cap
            reg["ocupacion_pct"] = 100.0 * ton / cap if cap > TOL_TON else None
            reg["costo_flete_M_CLP"] = sum(d.qV[u] * d.fV[u] * cunit * viajes_u[u] for u in d.V)
            return reg

        for t in d.T:
            gen = sum(d.gen[r, t] for r in d.R)
            alpha_t = d.alpha_t.get(t, d.alpha)

            rec = sum(X("x", o, c, t) for o in d.R for c in d.C) + sum(X("z", o, p, t) for o in d.R for p in d.P)
            val = sum(X("v", o, p, k, t) for o in d.R for p in d.P for k in d.K)
            mat = sum(X("v", o, p, k, t) for o in d.R for p in d.P for k in d.KM)

            stock_previo = (
                sum(X("s", o, c, anterior[t]) for o in d.R for c in d.C)
                if anterior[t] is not None else sum(d.s0[c] for c in d.C)
            )
            stock_final = sum(X("s", o, c, t) for o in d.R for c in d.C)
            residuo = rec + stock_previo - val - stock_final
            balances.append(dict(anio=t, recoleccion_ton=rec, inventario_inicial_ton=stock_previo,
                                 valorizacion_ton=val, inventario_final_ton=stock_final, residuo_balance_ton=residuo))

            ds8.append(dict(
                anio=t, generacion_ton=gen,
                meta_rec_ton=d.mr[t] * gen, recoleccion_ton=rec, pct_recoleccion=100.0 * rec / gen,
                deficit_rec_ton=X("dR", t),
                meta_val_ton=d.mv[t] * gen, valorizacion_ton=val, pct_valorizacion=100.0 * val / gen,
                deficit_val_ton=X("dV", t),
                valorizacion_material_ton=mat, valorizacion_energetica_ton=val - mat,
                pct_material_sobre_valorizado=100.0 * mat / val if val > TOL_TON else None,
                minimo_material_ton=alpha_t * val, holgura_material_ton=mat - alpha_t * val,
                deficit_cuotas_regionales_ton=sum(X("dQ", r, t) for r in d.R),
            ))

            for r in d.R:
                recogido = sum(X("x", r, c, t) for c in d.C) + sum(X("z", r, p, t) for p in d.P)
                cuotas.append(dict(anio=t, region_id=r, region=nr[r], disponible_ton=d.gen[r, t],
                                   cuota_minima_ton=d.qr[r] * gen, recoleccion_ton=recogido,
                                   deficit_ton=X("dQ", r, t)))

            for c in d.C:
                activo = X("y", c, t)
                apertura = activo - (X("y", c, anterior[t]) if anterior[t] is not None else 0)
                entrada = sum(X("x", o, c, t) for o in d.R)
                despacho = sum(X("f", o, c, p, t) for o in d.R for p in d.P)
                acopios.append(dict(
                    anio=t, centro_id=c, centro=nc[c], activo=round(activo),
                    cap_entrada_ton=d.aC[c] * activo, flujo_entrada_ton=entrada,
                    uso_entrada_pct=100.0 * entrada / (d.aC[c] * activo) if d.aC[c] * activo > TOL_TON else None,
                    cap_salida_ton=d.bC[c] * activo, flujo_salida_ton=despacho,
                    uso_salida_pct=100.0 * despacho / (d.bC[c] * activo) if d.bC[c] * activo > TOL_TON else None,
                    inv_max_ton=d.sM[c] * activo, inventario_final_ton=sum(X("s", o, c, t) for o in d.R),
                ))
                aperturas.append(dict(anio=t, tipo="Centro", instalacion=c, tecnologia="",
                                      activo=round(activo), apertura=round(apertura),
                                      inversion_M_CLP=d.iC[c] * apertura))

            for p in d.P:
                for k in d.K:
                    activo = X("w", p, k, t)
                    apertura = activo - (X("w", p, k, anterior[t]) if anterior[t] is not None else 0)
                    proc = sum(X("v", o, p, k, t) for o in d.R)
                    plantas.append(dict(
                        anio=t, planta_id=p, planta=np_[p], tecnologia_id=k, tecnologia=nk[k],
                        categoria_ds8="Material" if k in d.KM else "Energetica",
                        compatible=d.epk[p, k], activa=round(activo),
                        cap_efectiva_ton=d.aP[p, k] * activo, valorizado_ton=proc,
                        uso_capacidad_pct=100.0 * proc / (d.aP[p, k] * activo) if d.aP[p, k] * activo > TOL_TON else None,
                    ))
                    aperturas.append(dict(anio=t, tipo="Linea de planta", instalacion=p, tecnologia=k,
                                          activo=round(activo), apertura=round(apertura),
                                          inversion_M_CLP=d.iP[p, k] * apertura))

            # Flujos y viajes por tramo (se reporta todo arco con toneladas o viajes > 0)
            for r in d.R:
                for c in d.C:
                    ton, vj = X("x", r, c, t), {u: N("nx", r, c, u, t) for u in d.V}
                    if ton > TOL_TON or sum(vj.values()) > 0:
                        flujos.append(registro_flujo(t, "T1_Region_Acopio", r, nr[r], c, nc[c],
                                                     d.distT1[r, c], d.cT1[r, c], ton, vj, d.tauT1[r, c]))
            for c in d.C:
                for p in d.P:
                    ton = sum(X("f", o, c, p, t) for o in d.R)
                    vj = {u: N("nf", c, p, u, t) for u in d.V}
                    if ton > TOL_TON or sum(vj.values()) > 0:
                        flujos.append(registro_flujo(t, "T2_Acopio_Planta", c, nc[c], p, np_[p],
                                                     d.distT2[c, p], d.cT2[c, p], ton, vj, d.tauT2[c, p]))
            for r in d.R:
                for p in d.P:
                    ton, vj = X("z", r, p, t), {u: N("nz", r, p, u, t) for u in d.V}
                    if ton > TOL_TON or sum(vj.values()) > 0:
                        flujos.append(registro_flujo(t, "T3_Region_Planta_Directo", r, nr[r], p, np_[p],
                                                     d.distT3[r, p], d.cT3[r, p], ton, vj, d.tauT3[r, p]))

            # Resumen anual de viajes por tipo de camión
            ft = [fl for fl in flujos if fl["anio"] == t]
            fila_v = {"anio": t}
            for u in d.V:
                fila_v[f"viajes_{u}"] = sum(fl[f"viajes_{u}"] for fl in ft)
            fila_v["viajes_total"] = sum(fl["viajes_total"] for fl in ft)
            ton_mov = sum(fl["toneladas"] for fl in ft)
            cap_mov = sum(fl["capacidad_contratada_ton"] for fl in ft)
            fila_v["toneladas_transportadas"] = ton_mov
            fila_v["capacidad_contratada_ton"] = cap_mov
            fila_v["ocupacion_media_pct"] = 100.0 * ton_mov / cap_mov if cap_mov > TOL_TON else None
            fila_v["costo_flete_M_CLP"] = sum(fl["costo_flete_M_CLP"] for fl in ft)
            viajes_anual.append(fila_v)

            # Uso del tope de flota en regiones extremas (R13)
            for r in d.RE:
                centros_r = [c for c in d.C if d.region_centro[c] == r]
                for u in d.V:
                    if (r, u) not in d.tope:
                        continue
                    usados = (
                        sum(d.tauT1[r, c] * X("nx", r, c, u, t) for c in d.C)
                        + sum(d.tauT3[r, p] * X("nz", r, p, u, t) for p in d.P)
                        + sum(d.tauT2[c, p] * X("nf", c, p, u, t) for c in centros_r for p in d.P)
                    )
                    disp = d.tope[r, u] * d.dias
                    uso_flota.append(dict(anio=t, region_id=r, region=nr[r], camion=u,
                                          camiones_disponibles=d.tope[r, u],
                                          camion_dias_disponibles=disp, camion_dias_usados=usados,
                                          uso_pct=100.0 * usados / disp if disp > 0 else None))

        tablas.update({
            "Cumplimiento_Normativo_DS8": pd.DataFrame(ds8),
            "Cuotas_Regionales": pd.DataFrame(cuotas),
            "Capacidad_Centros_Acopio": pd.DataFrame(acopios),
            "Capacidad_Plantas_Tec": pd.DataFrame(plantas),
            "Matriz_Flujos_Viajes": pd.DataFrame(flujos),
            "Viajes_Flota_Anual": pd.DataFrame(viajes_anual),
            "Uso_Tope_Flota": pd.DataFrame(uso_flota),
            "Aperturas_Inversion": pd.DataFrame(aperturas),
            "Balance_Global": pd.DataFrame(balances),
        })

        m.write(str(salida / "solucion.sol"))

    estado = pd.DataFrame([info])
    estado.to_csv(salida / "resumen_solver.csv", index=False, encoding="utf-8-sig")

    escribir_xlsx(salida / "Resultados_Optimizacion_REP.xlsx", {"Estado_Gurobi": estado, **tablas})

    if hay_solucion:
        try:
            escribir_xlsx(salida / "Tablas_Procesadas_Informe.xlsx", tablas_informe(tablas, d))
        except Exception as e:
            print(f"Error al generar tablas procesadas: {e}")
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
    parser = argparse.ArgumentParser(description="Red REP-NFU ('Valora Más' 2026-2030) - Modelo MILP en Gurobi")
    parser.add_argument("-d", "--datos", type=str, default=None,
                        help=f"Ruta explícita al Excel de entrada (por defecto '{CARPETA_DATOS}/datos.xlsx').")
    parser.add_argument("--time-limit", type=float, default=TIME_LIMIT,
                        help=f"Límite de tiempo de Gurobi en segundos (por defecto: {TIME_LIMIT}).")
    parser.add_argument("--mip-gap", type=float, default=MIP_GAP,
                        help=f"Tolerancia de optimalidad relativa MIPGap (por defecto: {MIP_GAP}).")
    args = parser.parse_args()

    ruta = ruta_datos(explicita=args.datos)
    d = cargar_datos(ruta)
    print(
        f"Dataset cargado desde '{ruta.name}': "
        f"{len(d.R)} regiones; {len(d.C)} centros; {len(d.P)} plantas; {len(d.K)} tecnologías; "
        f"{len(d.V)} tipos de camión; tope de flota en {d.RE}; años {d.T}."
    )

    base = Path(__file__).resolve().parent
    raiz = base.parent if base.name in ["Script Python", "Script_Python"] else base
    salida = raiz / CARPETA_SALIDA / datetime.now().astimezone().strftime("ejecucion_%Y%m%d_%H%M%S_%f")
    salida.mkdir(parents=True, exist_ok=True)
    shutil.copy2(ruta, salida / "entrada_datos.xlsx")

    m = None
    try:
        m, variables, costos, anterior = construir_modelo(d, salida, time_limit=args.time_limit, mip_gap=args.mip_gap)
        print("\n--- ESTADÍSTICAS DEL MODELO ---")
        print(f"Variables totales    : {m.NumVars:,}")
        print(f"  Binarias           : {m.NumBinVars:,}")
        print(f"  Enteras            : {m.NumIntVars - m.NumBinVars:,}")
        print(f"  Continuas          : {m.NumVars - m.NumIntVars:,}")
        print(f"Restricciones totales: {m.NumConstrs:,}")

        historial = []
        m.optimize(progreso_callback(historial))
        info = guardar_resultados(m, d, variables, costos, anterior, salida, historial)
        return 0 if info["hay_solucion"] else 2
    finally:
        if m is not None:
            m.dispose()


if __name__ == "__main__":
    sys.exit(main())
