#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Análisis de sensibilidad (j) y caso base de impacto (k) — Grupo 52.

Usa el mismo modelo de main.py (no lo modifica). Cada escenario:
  1. Genera una copia de datos.xlsx con el parámetro cambiado en
     'Instancias Excel/Escenarios/<escenario>.xlsx' (trazabilidad de datos).
  2. Construye y resuelve el modelo de main.py (el caso base de impacto agrega
     las decisiones fijas de la práctica actual).
  3. Guarda los resultados completos en 'Analisis de Excels/Escenarios/<escenario>_<fecha>/'.
Al final (o con --solo-resumen) consolida todo en
'Analisis de Excels/Escenarios/Resumen_Escenarios.xlsx' y genera los gráficos.

Ejemplos:
  python escenarios.py --grupo prioridad --time-limit 600
  python escenarios.py --grupo todos
  python escenarios.py --escenario mu_x0.25
  python escenarios.py --solo-resumen
"""

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

import main
from xlsx_io import LibroXlsx, escribir_xlsx

BASE = Path(__file__).resolve().parent
RAIZ = BASE.parent
DIR_INST = RAIZ / main.CARPETA_DATOS / "Escenarios"
DIR_OUT = RAIZ / main.CARPETA_SALIDA / "Escenarios"
COLS_MU = ["Penalizacion_muR [M$/ton]", "Penalizacion_muV [M$/ton]", "Penalizacion_muQ [M$/ton]"]
REGIONES_EXTREMAS = ["R01", "R15", "R16"]


# =============================================================================
# DEFINICIÓN DE ESCENARIOS (modificaciones sobre las hojas de datos.xlsx)
# =============================================================================

def f_mu(factor):
    def aplicar(h):
        for c in COLS_MU:
            h["Metas_DS8"][c] = h["Metas_DS8"][c].astype(float) * factor
    return aplicar


def f_crecimiento(tasa, nivel=1.0):
    def aplicar(h):
        g = h["Regiones_Generacion"]
        cols = [c for c in g.columns if c.startswith("Gen_") and c.endswith("[ton]")]
        anios = sorted(int(c.split("_")[1].split(" ")[0]) for c in cols)
        t0 = anios[0]
        base0 = g[f"Gen_{t0} [ton]"].astype(float) * nivel
        es_total = g["ID_Region"].astype(str).str.upper() == "TOTAL"
        for t in anios:
            col = f"Gen_{t} [ton]"
            g[col] = (base0 * (1 + tasa) ** (t - t0)).round(1)
            g.loc[es_total, col] = round(float(g.loc[~es_total, col].sum()), 1)
    return aplicar


def f_sin_camion(camion):
    def aplicar(h):
        tf = h["Tope_Flota"]
        tf.loc[tf.ID_Region.isin(REGIONES_EXTREMAS), f"Camiones_{camion} [un]"] = 0
    return aplicar


def f_dias(factor):
    def aplicar(h):
        pf = h["Parametros_Flota"]
        fila = pf.Parametro == "Dias_Operativos_D"
        pf.loc[fila, "Valor"] = float(pf.loc[fila, "Valor"].iloc[0]) * factor
    return aplicar


def f_transporte(factor):
    def aplicar(h):
        ct = h["Costos_Transporte"]
        ct["Costo_Unitario [M$/ton]"] = ct["Costo_Unitario [M$/ton]"].astype(float) * factor
    return aplicar


def restricciones_practica_actual(m, var, d):
    """Caso base de impacto: red tradicional de licitación regional.
    Cada región opera su centro de acopio (y=1 en todos los centros y años) y
    todo el NFU pasa por acopio (sin envío directo región-planta, z=0)."""
    m.addConstrs((var["y"][c, t] == 1 for c in d.C for t in d.T), "BASE_centros")
    m.addConstrs((var["z"][r, p, t] == 0 for r in d.R for p in d.P for t in d.T), "BASE_sin_directo")


# nombre: (parámetro, etiqueta, valor numérico para graficar, función sobre datos, restricción extra)
ESCENARIOS = {
    "base_impacto":       ("Impacto", "Práctica actual: centro en cada región, sin envío directo", None, None, restricciones_practica_actual),
    "mu_x0.1":            ("Penalización μ", "μ × 0,10", 0.10, f_mu(0.10), None),
    "mu_x0.15":           ("Penalización μ", "μ × 0,15", 0.15, f_mu(0.15), None),
    "mu_x0.2":            ("Penalización μ", "μ × 0,20", 0.20, f_mu(0.20), None),
    "mu_x0.25":           ("Penalización μ", "μ × 0,25", 0.25, f_mu(0.25), None),
    "mu_x0.5":            ("Penalización μ", "μ × 0,50", 0.50, f_mu(0.50), None),
    "mu_x2":              ("Penalización μ", "μ × 2", 2.00, f_mu(2.0), None),
    "flota_sin25t":       ("Flota extrema", "Sin camión 25 t en R01/R15/R16", None, f_sin_camion("Camion_25t"), None),
    "flota_mitad_dias":   ("Flota extrema", "Días operativos D × 0,5", 0.50, f_dias(0.5), None),
    "gen_0pct":           ("Generación NFU", "Crecimiento 0%/año", 0.00, f_crecimiento(0.0), None),
    "gen_5pct":           ("Generación NFU", "Crecimiento 5%/año", 0.05, f_crecimiento(0.05), None),
    "gen_nivel_-10pct":   ("Generación NFU", "Nivel inicial −10%", None, f_crecimiento(0.025, 0.9), None),
    "transp_x0.7":        ("Tarifa transporte", "Tarifas × 0,7", 0.70, f_transporte(0.7), None),
    "transp_x1.3":        ("Tarifa transporte", "Tarifas × 1,3", 1.30, f_transporte(1.3), None),
    "transp_x2":          ("Tarifa transporte", "Tarifas × 2", 2.00, f_transporte(2.0), None),
    "transp_x3":          ("Tarifa transporte", "Tarifas × 3", 3.00, f_transporte(3.0), None),
}

GRUPOS = {
    "prioridad": ["base_impacto", "mu_x0.1", "mu_x0.25", "mu_x0.5",
                  "flota_sin25t", "flota_mitad_dias", "gen_0pct", "gen_5pct"],
    "impacto": ["base_impacto"],
    "mu": ["mu_x0.1", "mu_x0.25", "mu_x0.5", "mu_x2"],
    "flota": ["flota_sin25t", "flota_mitad_dias"],
    "gen": ["gen_0pct", "gen_5pct", "gen_nivel_-10pct"],
    "transporte": ["transp_x0.7", "transp_x1.3"],
    "extra": ["mu_x0.15", "mu_x0.2", "transp_x2", "transp_x3"],
    "todos": list(ESCENARIOS),
}


# =============================================================================
# EJECUCIÓN
# =============================================================================

def generar_excel(ruta_base, nombre, aplicar):
    libro = LibroXlsx(ruta_base)
    hojas = {s: libro.hoja(s) for s in libro.sheet_names}
    if aplicar is not None:
        aplicar(hojas)
    DIR_INST.mkdir(parents=True, exist_ok=True)
    destino = DIR_INST / f"{nombre}.xlsx"
    escribir_xlsx(destino, hojas)
    return destino


def correr(nombre, ruta_base, time_limit):
    parametro, etiqueta, _, aplicar, extra = ESCENARIOS[nombre]
    ruta = generar_excel(ruta_base, nombre, aplicar)
    d = main.cargar_datos(ruta)
    salida = DIR_OUT / f"{nombre}_{datetime.now().astimezone().strftime('%Y%m%d_%H%M%S')}"
    salida.mkdir(parents=True, exist_ok=True)
    shutil.copy2(ruta, salida / "entrada_datos.xlsx")
    print(f"\n{'#' * 78}\n# ESCENARIO {nombre}: {parametro} — {etiqueta}\n{'#' * 78}")
    m = None
    try:
        m, var, costos, ant = main.construir_modelo(d, salida, time_limit=time_limit)
        if extra is not None:
            extra(m, var, d)
        historial = []
        m.optimize(main.progreso_callback(historial))
        main.guardar_resultados(m, d, var, costos, ant, salida, historial)
    finally:
        if m is not None:
            m.dispose()


# =============================================================================
# RESUMEN Y GRÁFICOS
# =============================================================================

_PE = {}


def parametros_emisiones():
    """Lee 'Instancias Excel/Parametros_Emisiones.xlsx' (factor CO2e del diésel y rendimientos)."""
    if not _PE:
        ruta = RAIZ / main.CARPETA_DATOS / "Parametros_Emisiones.xlsx"
        if ruta.is_file():
            df = LibroXlsx(ruta).hoja("Parametros_Emisiones")
            _PE.update(dict(zip(df.Parametro.astype(str), pd.to_numeric(df.Valor))))
    return _PE


def metricas(carpeta):
    libro = LibroXlsx(carpeta / "Resultados_Optimizacion_REP.xlsx")
    est = libro.hoja("Estado_Gurobi").iloc[0]
    fila = dict(estado=est["estado"], gap_pct=est["gap_pct"], tiempo_s=est["tiempo_solver_s"],
                Z_MM=est["objetivo_MM_CLP"])
    if "Resumen_Costos_FO" not in libro.sheet_names:
        return fila
    c = libro.hoja("Resumen_Costos_FO").set_index("Componente")["Costo_MM_CLP"]
    centros = [k for k in c.index if k.startswith(("IC_C", "CF_C", "CT_1", "CT_2", "CA_S"))]
    fila.update(
        plantas_fijo_inv_MM=float(c["CF_P operacion fija plantas"] + c["IC_P inversion habilitacion lineas"]),
        procesamiento_MM=float(c["CP_V procesamiento"]),
        transporte_MM=float(c["CT_3 transporte region-planta"] + c["CT_1 transporte region-acopio"] + c["CT_2 transporte acopio-planta"]),
        centros_MM=float(sum(c[k] for k in centros if not k.startswith(("CT_1", "CT_2")))),
        penalizaciones_MM=float(c["CP_M penalizaciones deficit"]),
    )
    ds8 = libro.hoja("Cumplimiento_Normativo_DS8")
    fila.update(
        deficit_rec_ton=float(ds8.deficit_rec_ton.sum()),
        deficit_val_ton=float(ds8.deficit_val_ton.sum()),
        deficit_cuotas_ton=float(ds8.deficit_cuotas_regionales_ton.sum()),
        recoleccion_ton=float(ds8.recoleccion_ton.sum()),
        pct_material_2030=pd.to_numeric(ds8.pct_material_sobre_valorizado, errors="coerce").iloc[-1],
    )
    ap = libro.hoja("Aperturas_Inversion")
    ab = ap[ap.apertura == 1]
    fila["lineas_abiertas"] = "; ".join(f"{r.instalacion}-{r.tecnologia}({int(r.anio)})" for r in ab[ab.tipo != "Centro"].itertuples())
    fila["centros_abiertos"] = int((ab.tipo == "Centro").sum())
    v = libro.hoja("Viajes_Flota_Anual")
    for col in [c for c in v.columns if c.startswith("viajes_")]:
        fila[col] = int(v[col].sum())

    # Métricas no monetarias (impacto): km recorridos, ton-km, diésel y CO2e
    pe = parametros_emisiones()
    fl = libro.hoja("Matriz_Flujos_Viajes")
    if len(fl) and pe:
        km_total, litros = 0.0, 0.0
        for col in [c for c in fl.columns if c.startswith("viajes_Camion_")]:
            camion = col.replace("viajes_", "")
            km = float((fl[col].astype(float) * fl.distancia_km.astype(float)).sum()) * pe["Factor_Viaje_Redondo"]
            km_total += km
            litros += km / pe[f"Rendimiento_{camion}"]
        fila.update(
            km_camion=km_total,
            ton_km=float((fl.toneladas.astype(float) * fl.distancia_km.astype(float)).sum()),
            diesel_litros=litros,
            co2e_ton=litros * pe["Factor_Emision_Diesel"] / 1000.0,
        )
    if "Uso_Tope_Flota" in libro.sheet_names:
        u = libro.hoja("Uso_Tope_Flota")
        if len(u):
            fila["uso_max_tope_flota_pct"] = float(u.uso_pct.max())
    return fila


def tablas_informe(df):
    """Tablas 5, 6 y 7 del informe, con las mismas filas y columnas que se muestran."""
    def col(nombre):
        return df[nombre] if nombre in df else pd.Series([None] * len(df), index=df.index)

    def n_lineas(txt):
        return len([x for x in str(txt).split(";") if x.strip()]) if isinstance(txt, str) else 0

    t5 = pd.DataFrame({
        "Escenario": df.escenario,
        "Estado": df.estado,
        "Gap [%]": df.gap_pct,
        "Z [MM$]": df.Z_MM,
        "Delta Z [%]": col("delta_Z_pct"),
        "Lineas abiertas [un]": df.lineas_abiertas.map(n_lineas),
        "Lineas abiertas (anio)": df.lineas_abiertas,
        "Emisiones [t CO2e]": col("co2e_ton"),
    })
    t7 = pd.DataFrame({
        "Escenario": df.escenario,
        "Tiempo [s]": df.tiempo_s,
        "Recoleccion [t]": col("recoleccion_ton"),
        "Deficit recoleccion dR [t]": col("deficit_rec_ton"),
        "Deficit valorizacion dV [t]": col("deficit_val_ton"),
        "Deficit cuotas dQ [t]": col("deficit_cuotas_ton"),
        "Viajes 10t": col("viajes_Camion_10t"),
        "Viajes 15t": col("viajes_Camion_15t"),
        "Viajes 25t": col("viajes_Camion_25t"),
        "km [miles]": col("km_camion") / 1000.0,
        "Diesel [miles L]": col("diesel_litros") / 1000.0,
        "Uso max tope flota [%]": col("uso_max_tope_flota_pct"),
    })
    hojas = {"Tabla5_Sensibilidad": t5, "Tabla7_Anexo_Metricas": t7}

    if {"BASE (corrida oficial)", "base_impacto"} <= set(df.escenario):
        b = df[df.escenario == "base_impacto"].iloc[0]
        o = df[df.escenario == "BASE (corrida oficial)"].iloc[0]
        indicadores = [
            ("Costo total Z [MM$]", b.Z_MM, o.Z_MM),
            ("Centros de acopio [MM$]", b.centros_MM, o.centros_MM),
            ("Transporte [MM$]", b.transporte_MM, o.transporte_MM),
            ("Plantas, procesamiento y penalizacion [MM$]",
             b.plantas_fijo_inv_MM + b.procesamiento_MM + b.penalizaciones_MM,
             o.plantas_fijo_inv_MM + o.procesamiento_MM + o.penalizaciones_MM),
            ("Costo medio [M$/ton]", 1000 * b.Z_MM / b.recoleccion_ton, 1000 * o.Z_MM / o.recoleccion_ton),
            ("Viajes de camion [un]", b.viajes_total, o.viajes_total),
            ("Distancia recorrida [miles de km]", b.km_camion / 1000, o.km_camion / 1000),
            ("Diesel [miles de L]", b.diesel_litros / 1000, o.diesel_litros / 1000),
            ("Emisiones [t CO2e]", b.co2e_ton, o.co2e_ton),
        ]
        t6 = pd.DataFrame(indicadores, columns=["Indicador", "Caso base (red regional)", "Modelo"])
        t6["Diferencia"] = t6["Modelo"] - t6["Caso base (red regional)"]
        t6["Diferencia [%]"] = 100 * t6["Diferencia"] / t6["Caso base (red regional)"].where(t6["Caso base (red regional)"] != 0)
        hojas["Tabla6_Impacto"] = t6
    return hojas


def resumen():
    filas = []
    salida = RAIZ / main.CARPETA_SALIDA
    oficial = sorted(p for p in salida.glob("ejecucion_*") if (p / "Resultados_Optimizacion_REP.xlsx").is_file())
    if oficial:
        filas.append(dict(escenario="BASE (corrida oficial)", parametro="Base", descripcion="Instancia base",
                          carpeta=oficial[-1].name, **metricas(oficial[-1])))
    for nombre, (parametro, etiqueta, valor, _, _) in ESCENARIOS.items():
        corridas = sorted(p for p in DIR_OUT.glob(f"{nombre}_2*") if (p / "Resultados_Optimizacion_REP.xlsx").is_file())
        if corridas:  # se usa la corrida más reciente de cada escenario
            filas.append(dict(escenario=nombre, parametro=parametro, descripcion=etiqueta, valor=valor,
                              carpeta=corridas[-1].name, **metricas(corridas[-1])))
    if not filas:
        print("No hay corridas para resumir.")
        return
    df = pd.DataFrame(filas)
    if "BASE (corrida oficial)" in set(df.escenario):
        z0 = float(df.loc[df.escenario == "BASE (corrida oficial)", "Z_MM"].iloc[0])
        df["delta_Z_MM"] = df.Z_MM - z0
        df["delta_Z_pct"] = 100 * (df.Z_MM - z0) / z0
    DIR_OUT.mkdir(parents=True, exist_ok=True)
    escribir_xlsx(DIR_OUT / "Resumen_Escenarios.xlsx", {"Resumen": df, **tablas_informe(df)})
    graficos(df)
    print(df[["escenario", "estado", "gap_pct", "Z_MM"] + (["delta_Z_pct"] if "delta_Z_pct" in df else [])].to_string(index=False))
    print(f"\nResumen guardado en: {DIR_OUT / 'Resumen_Escenarios.xlsx'}")


def graficos(df):
    if "delta_Z_pct" not in df:
        return
    sens = df[~df.parametro.isin(["Base", "Impacto"])]
    if len(sens):
        fig, ax = plt.subplots(figsize=(10, 5))
        colores = {p: plt.cm.Set2(i) for i, p in enumerate(sens.parametro.unique())}
        ax.barh(sens.descripcion, sens.delta_Z_pct, color=[colores[p] for p in sens.parametro])
        ax.axvline(0, color="black", linewidth=0.8)
        for y, (val, est) in enumerate(zip(sens.delta_Z_pct, sens.estado)):
            ax.text(val, y, f" {val:+.1f}%" + ("" if est == "OPTIMAL" else "*"), va="center", fontsize=9)
        ax.set_xlabel("Variación de Z* respecto de la base [%]  (* = cortado por tiempo)")
        ax.set_title("Sensibilidad del costo total")
        ax.legend(handles=[plt.Rectangle((0, 0), 1, 1, color=c) for c in colores.values()],
                  labels=list(colores), loc="upper center", bbox_to_anchor=(0.5, -0.15),
                  ncol=len(colores), fontsize=8, frameon=False)
        ax.margins(x=0.15)
        ax.invert_yaxis()
        fig.savefig(DIR_OUT / "grafico_sensibilidad_Z.png", bbox_inches="tight")
        plt.close(fig)

    comp = ["plantas_fijo_inv_MM", "procesamiento_MM", "transporte_MM", "centros_MM", "penalizaciones_MM"]
    if all(c in df for c in comp):
        nombres = ["Plantas (fijo + inversión)", "Procesamiento", "Transporte", "Centros de acopio", "Penalizaciones"]
        sub = df.dropna(subset=comp)
        fig, ax = plt.subplots(figsize=(11, 5))
        izq = pd.Series(0.0, index=sub.index)
        for c, n, col in zip(comp, nombres, plt.cm.Set2.colors):
            ax.barh(sub.escenario, sub[c], left=izq, label=n, color=col)
            izq += sub[c]
        ax.set_xlabel("Costo total [MM$]")
        ax.set_title("Composición del costo por escenario")
        ax.legend(fontsize=8, loc="lower right")
        ax.invert_yaxis()
        fig.savefig(DIR_OUT / "grafico_composicion_costos.png", bbox_inches="tight")
        plt.close(fig)


# =============================================================================

def principal():
    ap = argparse.ArgumentParser(description="Escenarios de sensibilidad y caso base de impacto (usa el modelo de main.py).")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--grupo", choices=list(GRUPOS), help="Grupo de escenarios a correr.")
    g.add_argument("--escenario", choices=list(ESCENARIOS), help="Un escenario individual.")
    g.add_argument("--solo-resumen", action="store_true", help="Solo consolida las corridas existentes.")
    ap.add_argument("--time-limit", type=float, default=main.TIME_LIMIT, help="Límite de tiempo por escenario [s].")
    ap.add_argument("-d", "--datos", default=None, help="Excel base (por defecto Instancias Excel/datos.xlsx).")
    args = ap.parse_args()

    if not args.solo_resumen:
        ruta_base = main.ruta_datos(args.datos)
        lista = GRUPOS[args.grupo] if args.grupo else [args.escenario]
        print(f"Escenarios a correr ({len(lista)}): {lista} | límite {args.time_limit:.0f} s c/u")
        for nombre in lista:
            correr(nombre, ruta_base, args.time_limit)
    resumen()
    return 0


if __name__ == "__main__":
    sys.exit(principal())
