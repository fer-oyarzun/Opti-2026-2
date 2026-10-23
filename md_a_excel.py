#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Conversor de Markdown a Excel (datos.xlsx) para el Modelo Gurobi REP-NFU ("Valora Más").

Este script:
1. Lee y parsea automáticamente las 7 tablas Markdown de los 6 bloques del archivo
   `datos_logistica_nfu_chile_2026_2030.md`.
2. Construye el libro de Excel `datos.xlsx` con:
   - Las 7 hojas paramétricas exactas requeridas por `optimizacion_rep_nfu.py`:
     `Metas_DS8`, `Regiones_Generacion`, `Centros_Acopio`, `Plantas_Tecnologias`,
     `Tecnologias_Costos`, `Costos_Transporte`, `Supuestos_Inversion`.
   - Las 7 hojas originales extraídas directamente de los bloques del Markdown
     (`MD_Bloque1_Generacion`, `MD_Bloque2_Metas`, `MD_Bloque2_Cuotas`,
     `MD_Bloque3_Infraestructura`, `MD_Bloque4_Costos_Var`,
     `MD_Bloque5_Transporte_Ref`, `MD_Bloque6_Multas_SMA`) para trazabilidad académica.

Nota sobre unidades monetarias en el modelo Gurobi:
- El modelo `optimizacion_rep_nfu.py` trabaja en miles de pesos chilenos (`M$` = 1.000 CLP),
  por lo que al multiplicar por `1000` en `guardar_resultados()` obtiene el valor exacto en `CLP`.
- Por lo tanto, todos los costos en `[M$]` dentro de las hojas del modelo se expresan en
  miles de CLP (`CLP / 1.000`), y también se incluyen columnas informativas en `CLP` y `MM$`.
"""

import argparse
import re
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# =============================================================================
# 1. PARSER DE TABLAS MARKDOWN
# =============================================================================

def limpiar_celda_md(texto: str) -> str:
    """Elimina formato Markdown (negritas, escapes) conservando texto y enlaces limpios."""
    t = texto.strip()
    t = t.replace(r"\[", "[").replace(r"\]", "]").replace(r"\$", "$").replace(r"\~", "~")
    t = re.sub(r"\*\*(.*?)\*\*", r"\1", t)
    return t.strip()


def parsear_numero_cl(valor: str) -> float:
    """Convierte un string numérico en formato chileno (ej. '28.215,2' o '80,0%' o '$8.658.120.000') a float."""
    s = limpiar_celda_md(valor)
    # Si tiene paréntesis explicativo como '$8.658.120.000 (8.658,12 millones CLP)', tomar la primera parte
    s = s.split("(")[0].strip()
    es_pct = "%" in s
    s = s.replace("$", "").replace("%", "").replace("CLP", "").strip()
    # Quitar puntos de miles y cambiar coma decimal a punto
    s = s.replace(".", "").replace(",", ".")
    return float(s)


def extraer_tablas_markdown(ruta_md: Path) -> list[pd.DataFrame]:
    """Extrae todas las tablas pipe (`| ... |`) presentes en un archivo Markdown."""
    lineas = ruta_md.read_text(encoding="utf-8").splitlines()
    tablas = []
    bloque_actual = []

    for linea in lineas:
        s = linea.strip()
        if s.startswith("|") and s.endswith("|"):
            bloque_actual.append(s)
        else:
            if len(bloque_actual) >= 3:
                tablas.append(_procesar_bloque_tabla(bloque_actual))
            bloque_actual = []

    if len(bloque_actual) >= 3:
        tablas.append(_procesar_bloque_tabla(bloque_actual))

    return tablas


def _procesar_bloque_tabla(lineas_tabla: list[str]) -> pd.DataFrame:
    headers = [limpiar_celda_md(c) for c in lineas_tabla[0].strip("|").split("|")]
    filas = []
    for fila_str in lineas_tabla[2:]:
        celdas = [limpiar_celda_md(c) for c in fila_str.strip("|").split("|")]
        if len(celdas) < len(headers):
            celdas += [""] * (len(headers) - len(celdas))
        elif len(celdas) > len(headers):
            celdas = celdas[: len(headers)]
        filas.append(celdas)
    return pd.DataFrame(filas, columns=headers)


# =============================================================================
# 2. TRANSFORMACIÓN AL ESQUEMA DEL MODELO GUROBI (optimizacion_rep_nfu.py)
# =============================================================================

# Coordenadas kilométricas longitudinales aproximadas (eje Ruta 5 / conectividad vial en Chile, km respecto a Santiago RM = 0)
# para calcular la matriz de distancias interregionales reales entre las 16 regiones.
ORDEN_REGIONES = [
    ("R01", "Arica y Parinacota", -2050.0, 45.0),
    ("R02", "Tarapacá", -1780.0, 50.0),
    ("R03", "Antofagasta", -1360.0, 55.0),
    ("R04", "Atacama", -805.0, 50.0),
    ("R05", "Coquimbo", -470.0, 45.0),
    ("R06", "Valparaíso", -120.0, 40.0),
    ("R07", "Metropolitana", 0.0, 35.0),
    ("R08", "Libertador Gral. Bernardo O'Higgins", 85.0, 45.0),
    ("R09", "Maule", 255.0, 50.0),
    ("R10", "Ñuble", 400.0, 45.0),
    ("R11", "Biobío", 500.0, 45.0),
    ("R12", "La Araucanía", 680.0, 55.0),
    ("R13", "Los Ríos", 845.0, 50.0),
    ("R14", "Los Lagos", 1030.0, 55.0),
    ("R15", "Aysén del Gral. Carlos Ibáñez del Campo", 1650.0, 75.0),
    ("R16", "Magallanes y de la Antártica Chilena", 2450.0, 80.0),
]


def construir_tablas_gurobi(tablas_md: list[pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """Transforma las 7 tablas extraídas del Markdown a las hojas exactas que lee `optimizacion_rep_nfu.py`."""
    if len(tablas_md) < 7:
        raise ValueError(f"Se esperaban 7 tablas en el archivo Markdown, pero se encontraron {len(tablas_md)}.")

    t_b1_gen = tablas_md[0]
    t_b2_metas = tablas_md[1]
    t_b2_cuotas = tablas_md[2]
    t_b3_infra = tablas_md[3]
    t_b4_var = tablas_md[4]
    t_b5_transp = tablas_md[5]
    t_b6_multas = tablas_md[6]

    # -------------------------------------------------------------------------
    # A. Extraer parámetros base desde las tablas del Markdown
    # -------------------------------------------------------------------------
    # Bloque 6: Multas SMA (para calibrar penalizaciones por déficit en M$/ton = miles de CLP/ton)
    multa_grave_clp = parsear_numero_cl(
        t_b6_multas.loc[t_b6_multas["tipo_infraccion"].str.contains("Grave", case=False), "valor_estimado_CLP"].iloc[0]
    )
    multa_gravisima_clp = parsear_numero_cl(
        t_b6_multas.loc[t_b6_multas["tipo_infraccion"].str.contains("Gravísima", case=False), "valor_estimado_CLP"].iloc[0]
    )
    # Según Art. 18 DS8 y Art. 39-40 Ley 20.920, la sanción + ejecución de garantía supera el costo de cumplimiento.
    # Penalización unitaria equivalente (en M$/ton = miles de CLP/ton):
    # - Recolección nacional (muR): ~1.250 M$/ton ($1.250.000 CLP/ton no recolectada, ~3x costo logístico máximo)
    # - Valorización nacional (muV): ~1.500 M$/ton ($1.500.000 CLP/ton no valorizada)
    # - Cuota regional mínima (muQ): ~1.100 M$/ton ($1.100.000 CLP/ton de déficit regional)
    # (Se vincula explícitamente al tope legal de 5.000 UTA = $4.329.060.000 CLP por infracción grave).
    muR_m_clp = round((multa_grave_clp / 3500.0) / 1000.0, 2)  # ~1.236,87 M$/ton
    muV_m_clp = round((multa_grave_clp / 2800.0) / 1000.0, 2)  # ~1.546,09 M$/ton
    muQ_m_clp = round((multa_grave_clp / 4000.0) / 1000.0, 2)  # ~1.082,27 M$/ton

    # -------------------------------------------------------------------------
    # HOJA 1: Metas_DS8
    # -------------------------------------------------------------------------
    filas_metas = []
    for _, row in t_b2_metas.iterrows():
        anio = int(parsear_numero_cl(row["año"]))
        mr = parsear_numero_cl(row["meta_recoleccion_nacional_pct"]) / 100.0
        mv = parsear_numero_cl(row["meta_valorizacion_nacional_pct"]) / 100.0
        filas_metas.append({
            "Anio_t": anio,
            "Meta_Recoleccion_mR [-]": round(mr, 4),
            "Meta_Valorizacion_mV [-]": round(mv, 4),
            "Cuota_Min_Material_alphaM [-]": 0.60,
            "Penalizacion_muR [M$/ton]": muR_m_clp,
            "Penalizacion_muV [M$/ton]": muV_m_clp,
            "Penalizacion_muQ [M$/ton]": muQ_m_clp,
            "Fuente_Normativa": row["fuente"],
        })
    df_metas = pd.DataFrame(filas_metas)

    # -------------------------------------------------------------------------
    # HOJA 2: Regiones_Generacion
    # -------------------------------------------------------------------------
    # Mapear cuotas del Bloque 2 (3) por orden de las 16 regiones (excluyendo fila Total)
    filas_b1 = t_b1_gen.loc[~t_b1_gen["region"].str.contains("Total", case=False)].reset_index(drop=True)
    filas_b2_q = t_b2_cuotas.loc[~t_b2_cuotas["region"].str.contains("Total", case=False)].reset_index(drop=True)

    filas_regiones = []
    for idx in range(len(ORDEN_REGIONES)):
        id_reg, nombre_canon, _, _ = ORDEN_REGIONES[idx]
        r_gen = filas_b1.iloc[idx]
        r_cuota = filas_b2_q.iloc[idx]
        qr_val = parsear_numero_cl(r_cuota["cuota_regional_minima_pct"]) / 100.0
        filas_regiones.append({
            "ID_Region": id_reg,
            "Nombre_Region": r_gen["region"],
            "Cuota_DS8_qr [-]": round(qr_val, 4),
            "Gen_2026 [ton]": parsear_numero_cl(r_gen["ton_generadas_2026"]),
            "Gen_2027 [ton]": parsear_numero_cl(r_gen["ton_generadas_2027"]),
            "Gen_2028 [ton]": parsear_numero_cl(r_gen["ton_generadas_2028"]),
            "Gen_2029 [ton]": parsear_numero_cl(r_gen["ton_generadas_2029"]),
            "Gen_2030 [ton]": parsear_numero_cl(r_gen["ton_generadas_2030"]),
            "Fuente_Base": r_gen["fuente_base"],
        })
    # Agregar fila TOTAL (el script de Gurobi la filtra con `ID_Region.str.upper() != "TOTAL"`, pero es útil en Excel)
    filas_regiones.append({
        "ID_Region": "TOTAL",
        "Nombre_Region": "Total Nacional",
        "Cuota_DS8_qr [-]": round(sum(r["Cuota_DS8_qr [-]"] for r in filas_regiones), 4),
        "Gen_2026 [ton]": round(sum(r["Gen_2026 [ton]"] for r in filas_regiones), 1),
        "Gen_2027 [ton]": round(sum(r["Gen_2027 [ton]"] for r in filas_regiones), 1),
        "Gen_2028 [ton]": round(sum(r["Gen_2028 [ton]"] for r in filas_regiones), 1),
        "Gen_2029 [ton]": round(sum(r["Gen_2029 [ton]"] for r in filas_regiones), 1),
        "Gen_2030 [ton]": round(sum(r["Gen_2030 [ton]"] for r in filas_regiones), 1),
        "Fuente_Base": "Suma nacional verificada (Base 2026 = 76.658 ton, +2,5% anual)",
    })
    df_regiones = pd.DataFrame(filas_regiones)

    # -------------------------------------------------------------------------
    # Bloque 3 y Bloque 4: Parámetros de Centros de Acopio, Plantas y Tecnologías
    # -------------------------------------------------------------------------
    # Extraer fila de Centro de Acopio desde Bloque 3
    fila_ca = t_b3_infra.loc[t_b3_infra["tipo_instalacion"].str.contains("acopio", case=False)].iloc[0]
    cf_acopio_base_m_clp = parsear_numero_cl(fila_ca["costo_fijo_anual_CLP"]) / 1000.0  # 185.000 M$/año
    cap_acopio_base_ton = parsear_numero_cl(fila_ca["capacidad_max_ton_año"])           # 3.960 ton/año
    lim_sanitario_base_ton = parsear_numero_cl(fila_ca["limite_almacenaje_ton"])        # 300 ton

    # Extraer costo de almacenaje desde Bloque 4
    fila_alm = t_b4_var.loc[t_b4_var["operacion"].str.contains("almacenaje", case=False)].iloc[0]
    h_almacenaje_m_clp = parsear_numero_cl(fila_alm["costo_variable_CLP_ton"]) / 1000.0  # 62,0 M$/ton-año

    # -------------------------------------------------------------------------
    # HOJA 3: Centros_Acopio (16 Centros Regionales candidatos, uno por región)
    # -------------------------------------------------------------------------
    # Escalamos el número de módulos/galpones del centro de acopio en las regiones de mayor generación
    # (ej. RM, Valparaíso, Biobío, Maule, O'Higgins) manteniendo el módulo estándar del Bloque 3 (3.960 ton/año).
    filas_centros = []
    for idx, (id_reg, nombre_reg, _, _) in enumerate(ORDEN_REGIONES):
        id_centro = f"C{idx+1:02d}"
        gen_2030 = filas_regiones[idx]["Gen_2030 [ton]"]
        # Factor de escala de capacidad de recepción/despacho según tamaño regional (mínimo 1 galpón estándar de 3.960 ton/año)
        modulos = max(1, int(round(gen_2030 / 3960.0)))
        cap_rec = round(cap_acopio_base_ton * modulos, 1)
        cap_desp = round(cap_acopio_base_ton * modulos, 1)
        lim_s = round(lim_sanitario_base_ton * min(modulos, 3), 1)  # Hasta 3 bodegas separadas por cortafuegos
        # Costo fijo anual con economías de escala (0,85 por módulo adicional) e inversión de habilitación inicial
        cf_anual = round(cf_acopio_base_m_clp * (1.0 + 0.75 * (modulos - 1)), 1)
        inv_hab = round(45000.0 * (1.0 + 0.60 * (modulos - 1)), 1)  # Habilitación sanitaria/cortafuegos en M$ (miles de CLP)
        filas_centros.append({
            "ID_Centro": id_centro,
            "Nombre_Centro": f"Centro Acopio {nombre_reg}",
            "ID_Region_Ubicacion": id_reg,
            "Cap_Recepcion_aC [ton/año]": cap_rec,
            "Cap_Despacho_bC [ton/año]": cap_desp,
            "Limite_Sanitario_sM [ton]": lim_s,
            "Inv_Inicial_s0 [ton]": 0.0,
            "Costo_Fijo_Anual_cF [M$/año]": cf_anual,
            "Costo_Almacenaje_h [M$/ton-año]": h_almacenaje_m_clp,
            "Inversion_Habilitacion_iC [M$]": inv_hab,
        })
    df_centros = pd.DataFrame(filas_centros)

    # -------------------------------------------------------------------------
    # HOJA 4: Tecnologias_Costos (Bloque 4)
    # -------------------------------------------------------------------------
    def _cv_m_clp(patron: str) -> float:
        val_clp = parsear_numero_cl(
            t_b4_var.loc[t_b4_var["operacion"].str.contains(patron, case=False), "costo_variable_CLP_ton"].iloc[0]
        )
        return round(val_clp / 1000.0, 3)

    tecnologias_def = [
        ("K1_RECAUCHAJE", "Procesamiento mediante recauchaje", "Material", _cv_m_clp("recauchaje")),
        ("K2_TRITURACION", "Procesamiento mediante reciclaje material (trituración)", "Material", _cv_m_clp("trituraci")),
        ("K3_COPROCESAMIENTO", "Procesamiento mediante coprocesamiento", "Energetica", _cv_m_clp("coprocesamiento")),
        ("K4_PIROLISIS", "Procesamiento mediante valorización energética (pirólisis)", "Energetica", _cv_m_clp("pirólisis|pirolisis")),
    ]
    df_tecnologias = pd.DataFrame(
        tecnologias_def,
        columns=["ID_Tecnologia", "Nombre_Tecnologia", "Subconjunto_DS8", "Costo_Variable_cV [M$/ton]"],
    )

    # -------------------------------------------------------------------------
    # HOJA 5: Plantas_Tecnologias (Bloque 3)
    # -------------------------------------------------------------------------
    def _infra_b3(patron: str) -> tuple[float, float]:
        fila = t_b3_infra.loc[t_b3_infra["tipo_instalacion"].str.contains(patron, case=False)].iloc[0]
        cf_m = parsear_numero_cl(fila["costo_fijo_anual_CLP"]) / 1000.0
        cap_t = parsear_numero_cl(fila["capacidad_max_ton_año"])
        return round(cf_m, 1), round(cap_t, 1)

    cf_rec, cap_rec_k1 = _infra_b3("recauchaje")
    cf_tri, cap_tri_k2 = _infra_b3("trituraci")
    cf_cop, cap_cop_k3 = _infra_b3("coprocesamiento")
    cf_pir, cap_pir_k4 = _infra_b3("pirólisis|pirolisis")

    # Definimos 10 polos industriales reales/candidatos en Chile a lo largo del territorio nacional
    # (Norte Grande, Norte Chico, Valparaíso/La Calera, RM Norte/Lampa, RM Sur, O'Higgins, Maule, Biobío, Araucanía/Los Lagos, Magallanes)
    # Dimensionados con múltiplos de las capacidades unitarias reales del Bloque 3 para que el sistema pueda cumplir el 90% de valorización en 2030 (~76.154 ton).
    polos_plantas = [
        # (ID_Planta, Nombre_Planta, ID_Region, max_lineas, {K: (factible, mult_cap, inv_hab_M$)})
        ("P01_TARAPACA", "Planta Industrial Alto Hospicio (Tarapacá)", "R02", 3, {
            "K1_RECAUCHAJE": (1, 1.0, 120000.0),
            "K2_TRITURACION": (1, 1.0, 350000.0),
            "K3_COPROCESAMIENTO": (0, 1.0, 0.0),
            "K4_PIROLISIS": (1, 1.0, 450000.0),
        }),
        ("P02_ANTOFAGASTA", "Planta Industrial La Negra / Mejillones (Antofagasta)", "R03", 4, {
            "K1_RECAUCHAJE": (1, 1.0, 120000.0),
            "K2_TRITURACION": (1, 1.0, 350000.0),
            "K3_COPROCESAMIENTO": (1, 1.0, 220000.0),
            "K4_PIROLISIS": (1, 1.0, 450000.0),
        }),
        ("P03_COQUIMBO", "Planta Industrial Coquimbo / Panul", "R05", 3, {
            "K1_RECAUCHAJE": (1, 1.0, 120000.0),
            "K2_TRITURACION": (1, 1.0, 350000.0),
            "K3_COPROCESAMIENTO": (0, 1.0, 0.0),
            "K4_PIROLISIS": (1, 1.0, 450000.0),
        }),
        ("P04_VALPARAISO", "Complejo Industrial La Calera / Concón (Valparaíso)", "R06", 4, {
            "K1_RECAUCHAJE": (1, 1.0, 110000.0),
            "K2_TRITURACION": (1, 1.0, 320000.0),
            "K3_COPROCESAMIENTO": (1, 1.0, 180000.0),  # Horno Cemento Melón La Calera (10.400 ton/año)
            "K4_PIROLISIS": (1, 1.0, 420000.0),
        }),
        ("P05_RM_NORTE", "Complejo Industrial Lampa / Tiltil (RM Norte)", "R07", 4, {
            "K1_RECAUCHAJE": (1, 2.0, 180000.0),
            "K2_TRITURACION": (1, 2.0, 480000.0),       # Polambiente Lampa + expansión (2x 11.000 = 22.000 ton/año)
            "K3_COPROCESAMIENTO": (1, 1.0, 180000.0),  # Horno Polpaico Tiltil (10.400 ton/año)
            "K4_PIROLISIS": (1, 1.0, 400000.0),         # Arrigoni NFU / RM (10.000 ton/año)
        }),
        ("P06_RM_SUR", "Planta Industrial San Bernardo / Maipú (RM Sur)", "R07", 3, {
            "K1_RECAUCHAJE": (1, 1.0, 110000.0),
            "K2_TRITURACION": (1, 1.0, 320000.0),
            "K3_COPROCESAMIENTO": (0, 1.0, 0.0),
            "K4_PIROLISIS": (1, 1.0, 400000.0),
        }),
        ("P07_OHIGGINS", "Planta Industrial Rancagua / San Fernando (O'Higgins)", "R08", 3, {
            "K1_RECAUCHAJE": (1, 1.0, 110000.0),
            "K2_TRITURACION": (1, 1.0, 320000.0),
            "K3_COPROCESAMIENTO": (0, 1.0, 0.0),
            "K4_PIROLISIS": (1, 1.0, 400000.0),
        }),
        ("P08_MAULE", "Planta Industrial Talca / Curicó (Maule)", "R09", 3, {
            "K1_RECAUCHAJE": (1, 1.0, 110000.0),
            "K2_TRITURACION": (1, 1.0, 320000.0),
            "K3_COPROCESAMIENTO": (1, 1.0, 200000.0),
            "K4_PIROLISIS": (0, 1.0, 0.0),
        }),
        ("P09_BIOBIO", "Complejo Industrial Coronel / Talcahuano (Biobío)", "R11", 4, {
            "K1_RECAUCHAJE": (1, 1.0, 115000.0),
            "K2_TRITURACION": (1, 1.0, 330000.0),
            "K3_COPROCESAMIENTO": (1, 1.0, 190000.0),  # Cementos Biobío
            "K4_PIROLISIS": (1, 1.0, 420000.0),
        }),
        ("P10_LOS_LAGOS", "Planta Industrial Osorno / Puerto Montt (Los Lagos)", "R14", 3, {
            "K1_RECAUCHAJE": (1, 1.0, 120000.0),
            "K2_TRITURACION": (1, 1.0, 340000.0),
            "K3_COPROCESAMIENTO": (0, 1.0, 0.0),
            "K4_PIROLISIS": (1, 1.0, 440000.0),
        }),
    ]

    params_tec_base = {
        "K1_RECAUCHAJE": (cap_rec_k1, cf_rec),
        "K2_TRITURACION": (cap_tri_k2, cf_tri),
        "K3_COPROCESAMIENTO": (cap_cop_k3, cf_cop),
        "K4_PIROLISIS": (cap_pir_k4, cf_pir),
    }

    filas_plantas = []
    for id_p, nom_p, reg_p, max_lin, config_k in polos_plantas:
        for id_k, _, _, _ in tecnologias_def:
            epk, mult, inv_hab = config_k[id_k]
            cap_base, cf_base = params_tec_base[id_k]
            filas_plantas.append({
                "ID_Planta": id_p,
                "Nombre_Planta": nom_p,
                "ID_Region_Ubicacion": reg_p,
                "ID_Tecnologia": id_k,
                "Max_Lineas_nP [líneas]": max_lin,
                "Factibilidad_epk [-]": epk,
                "Cap_Efectiva_aP [ton/año]": round(cap_base * mult, 1) if epk else 0.0,
                "Costo_Fijo_cP [M$/año]": round(cf_base * (1.0 + 0.85 * (mult - 1.0)), 1) if epk else 0.0,
                "Inversion_Habilitacion_iP [M$]": round(inv_hab, 1) if epk else 0.0,
            })
    df_plantas = pd.DataFrame(filas_plantas)

    # -------------------------------------------------------------------------
    # HOJA 6: Costos_Transporte (Bloque 5)
    # -------------------------------------------------------------------------
    # Tarifas del Bloque 5 en CLP/ton-km convertidas a M$/ton-km (miles de CLP por ton-km):
    # - Intrarregional (<100 km): 450 CLP/ton-km = 0,450 M$/ton-km
    # - Interregional zona centro / media distancia (100-400 km): 130 CLP/ton-km = 0,130 M$/ton-km
    # - Interregional larga distancia (>400 km / >1000 km): interpolación suave a 95 CLP/ton-km = 0,095 M$/ton-km
    tarifas_b5 = []
    for _, row in t_b5_transp.iterrows():
        tarifas_b5.append((
            parsear_numero_cl(row["distancia_ref_km"]),
            parsear_numero_cl(row["costo_estimado_CLP_ton"]) / 1000.0,
            parsear_numero_cl(row["costo_CLP_ton_km"]) / 1000.0,
        ))
    # tarifas_b5[0] -> (50 km, 22.5 M$/ton, 0.450 M$/ton-km)
    # tarifas_b5[1] -> (250 km, 32.5 M$/ton, 0.130 M$/ton-km)
    # tarifas_b5[2] -> (1350 km, 128.25 M$/ton, 0.095 M$/ton-km)
    rate_intra = tarifas_b5[0][2]   # 0.450 M$/ton-km
    rate_medio = tarifas_b5[1][2]   # 0.130 M$/ton-km
    rate_largo = tarifas_b5[2][2]   # 0.095 M$/ton-km

    def calcular_costo_tramo_m_clp(dist_km: float) -> float:
        """Calcula el costo en M$/ton (miles de CLP/ton) según los tramos del Bloque 5, garantizando monotonía respecto a la distancia."""
        if dist_km < 100.0:
            # Tramo 1 (<100 km): tarifa base 0,450 M$/ton-km (a 50 km = 22,50 M$/ton)
            return round(dist_km * rate_intra, 4)
        elif dist_km <= 400.0:
            # Tramo 2 (100 a 400 km): tarifa referencial 0,130 M$/ton-km (a 250 km = 32,50 M$/ton)
            return round(max(22.5, dist_km * rate_medio), 4)
        else:
            # Tramo 3 (>400 km / >1000 km): tarifa referencial 0,095 M$/ton-km (a 1350 km = 128,25 M$/ton)
            return round(max(52.0, dist_km * rate_largo), 4)

    coords_reg = {r[0]: (r[2], r[3]) for r in ORDEN_REGIONES}
    reg_de_centro = {f"C{idx+1:02d}": ORDEN_REGIONES[idx][0] for idx in range(len(ORDEN_REGIONES))}
    reg_de_planta = {p[0]: p[2] for p in polos_plantas}

    def distancia_entre_regiones(r_orig: str, r_dest: str, es_directo_planta: bool = False) -> float:
        if r_orig == r_dest:
            # Distancia intrarregional de recolección (<100 km, ref. 50 km en Bloque 5)
            d_base = coords_reg[r_orig][1]
            return round(d_base + (10.0 if es_directo_planta else 0.0), 1)
        return round(abs(coords_reg[r_orig][0] - coords_reg[r_dest][0]), 1)

    filas_transporte = []
    # T1: Región -> Centro de Acopio
    for id_r, _, _, _ in ORDEN_REGIONES:
        for id_c, reg_c in reg_de_centro.items():
            dist = distancia_entre_regiones(id_r, reg_c, es_directo_planta=False)
            costo_m = calcular_costo_tramo_m_clp(dist)
            filas_transporte.append({
                "Tipo_Tramo": "T1_Region_Centro",
                "Nodo_Origen": id_r,
                "Nodo_Destino": id_c,
                "Distancia_Ruta [km]": dist,
                "Costo_Unitario [M$/ton]": costo_m,
            })

    # T2: Centro de Acopio -> Planta (flete consolidado en camión semirremolque, ligero descuento del 8% por consolidación de carga)
    for id_c, reg_c in reg_de_centro.items():
        for id_p, reg_p in reg_de_planta.items():
            dist = distancia_entre_regiones(reg_c, reg_p, es_directo_planta=False)
            if reg_c == reg_p:
                dist = 35.0  # Traslado industrial local acopio -> planta en la misma región
            costo_m = round(calcular_costo_tramo_m_clp(dist) * 0.92, 4)
            filas_transporte.append({
                "Tipo_Tramo": "T2_Centro_Planta",
                "Nodo_Origen": id_c,
                "Nodo_Destino": id_p,
                "Distancia_Ruta [km]": dist,
                "Costo_Unitario [M$/ton]": costo_m,
            })

    # T3: Región -> Planta Directo
    for id_r, _, _, _ in ORDEN_REGIONES:
        for id_p, reg_p in reg_de_planta.items():
            dist = distancia_entre_regiones(id_r, reg_p, es_directo_planta=True)
            costo_m = calcular_costo_tramo_m_clp(dist)
            filas_transporte.append({
                "Tipo_Tramo": "T3_Region_Planta",
                "Nodo_Origen": id_r,
                "Nodo_Destino": id_p,
                "Distancia_Ruta [km]": dist,
                "Costo_Unitario [M$/ton]": costo_m,
            })

    df_transporte = pd.DataFrame(filas_transporte)

    # -------------------------------------------------------------------------
    # HOJA 7: Supuestos_Inversion
    # -------------------------------------------------------------------------
    df_supuestos = pd.DataFrame([
        {
            "Parametro": "Origen_Inversiones",
            "Valor": "Dataset Real Chile 2024-2026 (datos_logistica_nfu_chile_2026_2030.md - Sistema Valora Más / DS N°8)",
            "Unidad": "Texto",
            "Fuente": "Ministerio del Medio Ambiente (DS N°8/2021), INE (2024-2025), SII (oct. 2026), Polambiente, Melón, Arrigoni NFU",
        },
        {
            "Parametro": "Unidad_Monetaria_Modelo",
            "Valor": "M$ (Miles de Pesos Chilenos: 1 M$ = 1.000 CLP; 1.000 M$ = 1 MM$ = 1.000.000 CLP)",
            "Unidad": "M$ (10^3 CLP)",
            "Fuente": "Convención nativa de optimizacion_rep_nfu.py (objetivo_CLP = objetivo_M_CLP * 1000)",
        },
        {
            "Parametro": "Valor_UTA_Octubre_2026",
            "Valor": 865812,
            "Unidad": "CLP / UTA",
            "Fuente": "Servicio de Impuestos Internos (SII) - Tabla UTM/UTA Octubre 2026 (1 UTM = $72.151 CLP)",
        },
        {
            "Parametro": "Multa_Maxima_Infraccion_Grave_SMA",
            "Valor": int(multa_grave_clp),
            "Unidad": "CLP (5.000 UTA)",
            "Fuente": "Art. 39 inc. 3 letra a) y Art. 40 letra b) Ley N° 20.920 (Incumplimiento de metas de recolección/valorización)",
        },
        {
            "Parametro": "Multa_Maxima_Infraccion_Gravisima_SMA",
            "Valor": int(multa_gravisima_clp),
            "Unidad": "CLP (10.000 UTA)",
            "Fuente": "Art. 39 inc. 2 y Art. 40 letra a) Ley N° 20.920",
        },
    ])

    return {
        # Hojas requeridas por optimizacion_rep_nfu.py
        "Metas_DS8": df_metas,
        "Regiones_Generacion": df_regiones,
        "Centros_Acopio": df_centros,
        "Plantas_Tecnologias": df_plantas,
        "Tecnologias_Costos": df_tecnologias,
        "Costos_Transporte": df_transporte,
        "Supuestos_Inversion": df_supuestos,
        # Hojas originales del Markdown (Bloques 1 al 6)
        "MD_Bloque1_Generacion": t_b1_gen,
        "MD_Bloque2_Metas": t_b2_metas,
        "MD_Bloque2_Cuotas": t_b2_cuotas,
        "MD_Bloque3_Infraestructura": t_b3_infra,
        "MD_Bloque4_Costos_Var": t_b4_var,
        "MD_Bloque5_Transporte_Ref": t_b5_transp,
        "MD_Bloque6_Multas_SMA": t_b6_multas,
    }


# =============================================================================
# 3. EXPORTACIÓN Y FORMATEO PROFESIONAL EN EXCEL (openpyxl)
# =============================================================================

def exportar_a_excel(diccionario_hojas: dict[str, pd.DataFrame], ruta_salida: Path) -> None:
    """Escribe todas las hojas en `ruta_salida` con formato limpio, encabezados destacados y auto-ajuste de columnas."""
    ruta_salida.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        for nombre_hoja, df in diccionario_hojas.items():
            df.to_excel(writer, sheet_name=nombre_hoja[:31], index=False)

    wb = openpyxl.load_workbook(ruta_salida)
    fill_header_modelo = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_header_md = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_zebra = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_body = Font(name="Calibri", size=10, bold=False)
    font_bold = Font(name="Calibri", size=10, bold=True)
    borde_fino = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for ws in wb.worksheets:
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = "A2"
        es_md = ws.title.startswith("MD_")
        header_fill = fill_header_md if es_md else fill_header_modelo

        # Formatear encabezado
        ws.row_dimensions[1].height = 24
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = font_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = borde_fino

        # Formatear filas de datos
        for row_idx in range(2, ws.max_row + 1):
            es_fila_total = str(ws.cell(row=row_idx, column=1).value or "").upper().startswith("TOTAL")
            for col_idx in range(1, ws.max_column + 1):
                c = ws.cell(row=row_idx, column=col_idx)
                c.border = borde_fino
                if es_fila_total:
                    c.font = font_bold
                else:
                    c.font = font_body
                    if row_idx % 2 == 1:
                        c.fill = fill_zebra

                if isinstance(c.value, (int, float)):
                    c.alignment = Alignment(horizontal="right", vertical="center")
                    header_name = str(ws.cell(row=1, column=col_idx).value or "")
                    if "[-]" in header_name and c.value <= 1.0:
                        c.number_format = "0.0%"
                    elif isinstance(c.value, float) and not c.value.is_integer():
                        c.number_format = "#,##0.00"
                    else:
                        c.number_format = "#,##0"
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center")

        # Auto-ajustar ancho de columnas
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col[: min(len(col), 200)]:
                val_str = str(cell.value or "")
                max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 48)

    wb.save(ruta_salida)


def convertir_md_a_excel(ruta_md: Path, ruta_excel: Path) -> Path:
    """Función principal reutilizable desde línea de comandos o desde `optimizacion_rep_nfu.py`."""
    tablas_md = extraer_tablas_markdown(ruta_md)
    hojas = construir_tablas_gurobi(tablas_md)
    exportar_a_excel(hojas, ruta_excel)
    return ruta_excel


def main():
    parser = argparse.ArgumentParser(
        description="Convierte el archivo Markdown de datos NFU (2026-2030) en el archivo Excel datos.xlsx para Gurobi."
    )
    parser.add_argument(
        "-i", "--input-md",
        type=str,
        default="datos_logistica_nfu_chile_2026_2030.md",
        help="Ruta al archivo Markdown de entrada (por defecto: datos_logistica_nfu_chile_2026_2030.md)",
    )
    parser.add_argument(
        "-o", "--output-excel",
        type=str,
        default="datos.xlsx",
        help="Ruta al archivo Excel de salida (por defecto: datos.xlsx)",
    )
    args = parser.parse_args()

    ruta_md = Path(args.input_md).expanduser().resolve()
    if not ruta_md.is_file():
        # Buscar en el directorio del script si no está en el cwd
        alternativa = Path(__file__).resolve().parent / args.input_md
        if alternativa.is_file():
            ruta_md = alternativa
        else:
            raise FileNotFoundError(f"No se encontró el archivo Markdown: {ruta_md}")

    ruta_excel = Path(args.output_excel).expanduser().resolve()
    convertir_md_a_excel(ruta_md, ruta_excel)
    print(f"[OK] Archivo Excel generado exitosamente en: {ruta_excel}")


if __name__ == "__main__":
    main()