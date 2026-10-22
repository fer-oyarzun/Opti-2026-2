from pathlib import Path
import pandas as pd
from gurobipy import GRB, Model, quicksum




# -- Cargar datos --
BASE_DIR = Path(__file__).resolve().parent
for cand in [BASE_DIR / "Instancias_Excel" / "datos.xlsx", BASE_DIR.parent / "Instancias_Excel" / "datos.xlsx", BASE_DIR / "datos.xlsx"]:
    if cand.exists():
        RUTA_EXCEL = cand
        break


RUTA_EXCEL = "dataset/datos.xlsx"
#creo q era asi, no era asi
regiones = pd.read_excel(RUTA_EXCEL, sheet_name="Regiones_Generacion")
regiones = regiones[regiones["ID_Region"].notna() & (regiones["ID_Region"] != "TOTAL")].set_index("ID_Region")

metas = pd.read_excel(RUTA_EXCEL, sheet_name="Metas_DS8").set_index("Anio_t")
centros = pd.read_excel(RUTA_EXCEL, sheet_name="Centros_Acopio").set_index("ID_Centro")
plantas_tec = pd.read_excel(RUTA_EXCEL, sheet_name="Plantas_Tecnologias")
tecnologias = pd.read_excel(RUTA_EXCEL, sheet_name="Tecnologias_Costos").set_index("ID_Tecnologia")
transporte = pd.read_excel(RUTA_EXCEL, sheet_name="Costos_Transporte")

# -- Conjuntos --
T_ = [int(t) for t in metas.index]
R_ = list(regiones.index)
C_ = list(centros.index)
P_ = sorted(plantas_tec["ID_Planta"].unique().tolist())
K_ = list(tecnologias.index)
KM_ = list(tecnologias.loc[tecnologias["Subconjunto_DS8"] == "Material"].index)
KE_ = list(tecnologias.loc[tecnologias["Subconjunto_DS8"] == "Energetica"].index)

# -- Parámetros --
## Demanda, generación regional y cuotas DS8
cuota_reg = {r: float(regiones.loc[r, "Cuota_DS8_qr [-]"]) for r in R_}
gen_nfu = {(r, t): float(regiones.loc[r, f"Gen_{t} [ton]"]) for r in R_ for t in T_}

meta_rec = {t: float(metas.loc[t, "Meta_Recoleccion_mR [-]"]) for t in T_}
meta_val = {t: float(metas.loc[t, "Meta_Valorizacion_mV [-]"]) for t in T_}
alpha_mat = float(metas["Cuota_Min_Material_alphaM [-]"].iloc[0])

multa_rec = {t: float(metas.loc[t, "Penalizacion_muR [M$/ton]"]) for t in T_}
multa_val = {t: float(metas.loc[t, "Penalizacion_muV [M$/ton]"]) for t in T_}
multa_reg = {(r, t): float(metas.loc[t, "Penalizacion_muQ [M$/ton]"]) for r in R_ for t in T_}

## Centros de acopio (capacidad recepción entrada, despacho salida e inventario sanitario)
cap_rec_ca = {c: float(centros.loc[c, "Cap_Recepcion_aC [ton/año]"]) for c in C_}
cap_des_ca = {c: float(centros.loc[c, "Cap_Despacho_bC [ton/año]"]) for c in C_}
inv_max_ca = {c: float(centros.loc[c, "Limite_Sanitario_sM [ton]"]) for c in C_}
inv_ini_ca = {c: float(centros.loc[c, "Inv_Inicial_s0 [ton]"]) for c in C_}
cfijo_ca = {(c, t): float(centros.loc[c, "Costo_Fijo_Anual_cF [M$/año]"]) for c in C_ for t in T_}
costo_alm = {c: float(centros.loc[c, "Costo_Almacenaje_h [M$/ton-año]"]) for c in C_}

## Plantas de valorización y tecnologías
max_lin_pl = {str(r["ID_Planta"]): int(r["Max_Lineas_nP [líneas]"]) for _, r in plantas_tec.iterrows()}
fact_pk = {(str(r["ID_Planta"]), str(r["ID_Tecnologia"])): int(r["Factibilidad_epk [-]"]) for _, r in plantas_tec.iterrows()}
cap_ef_pk = {(str(r["ID_Planta"]), str(r["ID_Tecnologia"])): float(r["Cap_Efectiva_aP [ton/año]"]) for _, r in plantas_tec.iterrows()}
cfijo_pl = {(str(r["ID_Planta"]), str(r["ID_Tecnologia"]), t): float(r["Costo_Fijo_cP [M$/año]"]) for _, r in plantas_tec.iterrows() for t in T_}
cvar_tec = {k: float(tecnologias.loc[k, "Costo_Variable_cV [M$/ton]"]) for k in K_}

## Costos y distancias de transporte en camión pesado por tramo
tr_t1 = transporte.loc[transporte["Tipo_Tramo"] == "T1_Region_Centro"]
tr_t2 = transporte.loc[transporte["Tipo_Tramo"] == "T2_Centro_Planta"]
tr_t3 = transporte.loc[transporte["Tipo_Tramo"] == "T3_Region_Planta"]

c_t1 = {(str(r["Nodo_Origen"]), str(r["Nodo_Destino"])): float(r["Costo_Unitario [M$/ton]"]) for _, r in tr_t1.iterrows()}
c_t2 = {(str(r["Nodo_Origen"]), str(r["Nodo_Destino"])): float(r["Costo_Unitario [M$/ton]"]) for _, r in tr_t2.iterrows()}
c_t3 = {(str(r["Nodo_Origen"]), str(r["Nodo_Destino"])): float(r["Costo_Unitario [M$/ton]"]) for _, r in tr_t3.iterrows()}

dist_t1 = {(str(r["Nodo_Origen"]), str(r["Nodo_Destino"])): float(r["Distancia_Ruta [km]"]) for _, r in tr_t1.iterrows()}
dist_t2 = {(str(r["Nodo_Origen"]), str(r["Nodo_Destino"])): float(r["Distancia_Ruta [km]"]) for _, r in tr_t2.iterrows()}
dist_t3 = {(str(r["Nodo_Origen"]), str(r["Nodo_Destino"])): float(r["Distancia_Ruta [km]"]) for _, r in tr_t3.iterrows()}

# -- Modelo --
modelo = Model("Optimizacion_REP_NFU")
modelo.setParam("TimeLimit", 1800)
modelo.setParam("MIPGap", 0.0001)

# var binarias
y = modelo.addVars(C_, T_, vtype=GRB.BINARY, name="y")
w = modelo.addVars(P_, K_, T_, vtype=GRB.BINARY, name="w")

# var continuas de flujo, procesamiento e inventario
x = modelo.addVars(R_, C_, T_, vtype=GRB.CONTINUOUS, name="x", lb=0)
z = modelo.addVars(R_, P_, T_, vtype=GRB.CONTINUOUS, name="z", lb=0)
f = modelo.addVars(C_, P_, T_, vtype=GRB.CONTINUOUS, name="f", lb=0)
v = modelo.addVars(P_, K_, T_, vtype=GRB.CONTINUOUS, name="v", lb=0)
s = modelo.addVars(C_, T_, vtype=GRB.CONTINUOUS, name="s", lb=0)

# var continuas de holgura (déficit normativo)
dR = modelo.addVars(T_, vtype=GRB.CONTINUOUS, name="dR", lb=0)
dV = modelo.addVars(T_, vtype=GRB.CONTINUOUS, name="dV", lb=0)
dQ = modelo.addVars(R_, T_, vtype=GRB.CONTINUOUS, name="dQ", lb=0)

modelo.update()

## Funcion objetivo
CF_C = quicksum(cfijo_ca[c, t] * y[c, t] for c in C_ for t in T_)
CF_P = quicksum(cfijo_pl[p, k, t] * w[p, k, t] for p in P_ for k in K_ for t in T_)
CT_1 = quicksum(c_t1[r, c] * x[r, c, t] for r in R_ for c in C_ for t in T_)
CT_2 = quicksum(c_t2[c, p] * f[c, p, t] for c in C_ for p in P_ for t in T_)
CT_3 = quicksum(c_t3[r, p] * z[r, p, t] for r in R_ for p in P_ for t in T_)
CP_V = quicksum(cvar_tec[k] * v[p, k, t] for p in P_ for k in K_ for t in T_)
CA_S = quicksum(costo_alm[c] * s[c, t] for c in C_ for t in T_)
CP_M = quicksum(multa_rec[t] * dR[t] + multa_val[t] * dV[t] + quicksum(multa_reg[r, t] * dQ[r, t] for r in R_) for t in T_)

modelo.setObjective(CF_C + CF_P + CT_1 + CT_2 + CT_3 + CP_V + CA_S + CP_M, GRB.MINIMIZE)

## -- Restricciones --

modelo.addConstrs((quicksum(x[r, c, t] for c in C_) + quicksum(z[r, p, t] for p in P_) <= gen_nfu[r, t]
                   for r in R_ for t in T_), name="disponibilidad_regional_nfu")

modelo.addConstrs((quicksum(x[r, c, t] for r in R_) <= cap_rec_ca[c] * y[c, t]
                   for c in C_ for t in T_), name="capacidad_recepcion_acopio")

modelo.addConstrs((quicksum(f[c, p, t] for p in P_) <= cap_des_ca[c] * y[c, t]
                   for c in C_ for t in T_), name="capacidad_despacho_acopio")

modelo.addConstrs((s[c, t] <= inv_max_ca[c] * y[c, t]
                   for c in C_ for t in T_), name="limite_sanitario_inventario")

modelo.addConstrs((s[c, T_[0]] == inv_ini_ca[c] + quicksum(x[r, c, T_[0]] for r in R_)
                   - quicksum(f[c, p, T_[0]] for p in P_)
                   for c in C_), name="balance_inventario_anio_uno")

modelo.addConstrs((s[c, T_[idx]] == s[c, T_[idx - 1]] + quicksum(x[r, c, T_[idx]] for r in R_)
                   - quicksum(f[c, p, T_[idx]] for p in P_)
                   for c in C_ for idx in range(1, len(T_))), name="balance_inventario_anios_siguientes")

modelo.addConstrs((quicksum(v[p, k, t] for k in K_) == quicksum(f[c, p, t] for c in C_) + quicksum(z[r, p, t] for r in R_)
                   for p in P_ for t in T_), name="balance_procesamiento_planta")

modelo.addConstrs((v[p, k, t] <= cap_ef_pk[p, k] * w[p, k, t]
                   for p in P_ for k in K_ for t in T_), name="capacidad_efectiva_planta_tecnologia")

modelo.addConstrs((w[p, k, t] <= fact_pk[p, k]
                   for p in P_ for k in K_ for t in T_), name="factibilidad_tecnica_planta")

modelo.addConstrs((quicksum(w[p, k, t] for k in K_) <= max_lin_pl[p]
                   for p in P_ for t in T_), name="maximo_tecnologias_por_planta")

modelo.addConstrs((quicksum(x[r, c, t] for r in R_ for c in C_) + quicksum(z[r, p, t] for r in R_ for p in P_) + dR[t]
                   >= meta_rec[t] * quicksum(gen_nfu[r, t] for r in R_)
                   for t in T_), name="meta_nacional_recoleccion")

modelo.addConstrs((quicksum(v[p, k, t] for p in P_ for k in K_) + dV[t]
                   >= meta_val[t] * quicksum(gen_nfu[r, t] for r in R_)
                   for t in T_), name="meta_nacional_valorizacion")

modelo.addConstrs(((1.0 - alpha_mat) * quicksum(v[p, k, t] for p in P_ for k in KM_)
                   - alpha_mat * quicksum(v[p, k, t] for p in P_ for k in KE_) >= 0.0
                   for t in T_), name="cuota_minima_reciclaje_material_ds8")

modelo.addConstrs((quicksum(x[r, c, t] for c in C_) + quicksum(z[r, p, t] for p in P_) + dQ[r, t]
                   >= cuota_reg[r] * quicksum(gen_nfu[rp, t] for rp in R_)
                   for r in R_ for t in T_), name="cuota_regional_recoleccion")

modelo.addConstrs((y[c, T_[idx]] >= y[c, T_[idx - 1]]
                   for c in C_ for idx in range(1, len(T_))), name="permanencia_centros_acopio")

modelo.addConstrs((w[p, k, T_[idx]] >= w[p, k, T_[idx - 1]]
                   for p in P_ for k in K_ for idx in range(1, len(T_))), name="permanencia_plantas_tecnologia")

modelo.optimize()

if modelo.status in (GRB.INFEASIBLE, GRB.INF_OR_UNBD):
    modelo.computeIIS()
    modelo.write("modelo_infactible.ilp")
    raise RuntimeError("Modelo infactible; revisar modelo_infactible.ilp")

# resultados
print("\n" + "=" * 70)
print(f" Valor objetivo: {modelo.ObjVal:,.2f} M$ ({modelo.ObjVal/1000:,.2f} MM CLP)    GAP: {modelo.MIPGap*100:.2f}%")
print(f" Tiempo: {modelo.Runtime:.2f}s    Variables: {modelo.NumVars:,}    Restricciones: {modelo.NumConstrs:,}")
print("=" * 70)

# diccionarios nombres
nom_r = regiones["Nombre_Region"].to_dict()
nom_c = centros["Nombre_Centro"].to_dict()
nom_p = plantas_tec.groupby("ID_Planta")["Nombre_Planta"].first().to_dict()
nom_k = tecnologias["Nombre_Tecnologia"].to_dict()

# costos
c_fijo_c = sum(cfijo_ca[c, t] * y[c, t].X for c in C_ for t in T_)
c_fijo_p = sum(cfijo_pl[p, k, t] * w[p, k, t].X for p in P_ for k in K_ for t in T_)
c_tr_1 = sum(c_t1[r, c] * x[r, c, t].X for r in R_ for c in C_ for t in T_)
c_tr_2 = sum(c_t2[c, p] * f[c, p, t].X for c in C_ for p in P_ for t in T_)
c_tr_3 = sum(c_t3[r, p] * z[r, p, t].X for r in R_ for p in P_ for t in T_)
c_proc = sum(cvar_tec[k] * v[p, k, t].X for p in P_ for k in K_ for t in T_)
c_alma = sum(costo_alm[c] * s[c, t].X for c in C_ for t in T_)
c_mult = sum(multa_rec[t] * dR[t].X + multa_val[t] * dV[t].X + sum(multa_reg[r, t] * dQ[r, t].X for r in R_) for t in T_)
total = c_fijo_c + c_fijo_p + c_tr_1 + c_tr_2 + c_tr_3 + c_proc + c_alma + c_mult

print(f"\nDesglose de costos (total = {total:,.2f} M$ / {total/1000:,.2f} MM CLP):")
print(f"  CP_V procesamiento plantas: {c_proc:>14,.2f} M$  ({c_proc/total*100:5.1f}%)")
print(f"  CF_P costo fijo plantas:    {c_fijo_p:>14,.2f} M$  ({c_fijo_p/total*100:5.1f}%)")
print(f"  CT_3 transporte directo:    {c_tr_3:>14,.2f} M$  ({c_tr_3/total*100:5.1f}%)")
print(f"  CT_1 transporte a acopio:   {c_tr_1:>14,.2f} M$  ({c_tr_1/total*100:5.1f}%)")
print(f"  CT_2 transporte troncal:    {c_tr_2:>14,.2f} M$  ({c_tr_2/total*100:5.1f}%)")
print(f"  CF_C costo fijo acopios:    {c_fijo_c:>14,.2f} M$  ({c_fijo_c/total*100:5.1f}%)")
print(f"  CA_S almacenaje sanitario:  {c_alma:>14,.2f} M$  ({c_alma/total*100:5.1f}%)")
print(f"  CP_M multas deficit DS8:    {c_mult:>14,.2f} M$  ({c_mult/total*100:5.1f}%)")

print("\nOperación acumulada por centro de acopio (2026-2030):")
for c in C_:
    anios_act = [t for t in T_ if y[c, t].X > 0.5]
    ap = anios_act[0] if anios_act else "-"
    flujo_in = sum(x[r, c, t].X for r in R_ for t in T_)
    flujo_out = sum(f[c, p, t].X for p in P_ for t in T_)
    inv_fin = s[c, T_[-1]].X
    print(f"  {c} ({nom_c[c]:28s}): apertura={ap} | entrada={flujo_in:8,.1f} t | salida={flujo_out:8,.1f} t | stock_2030={inv_fin:6,.1f} t")

print("\nCumplimiento normativo DS8 por año:")
for t in T_:
    gen_t = sum(gen_nfu[r, t] for r in R_)
    rec_t = sum(x[r, c, t].X for r in R_ for c in C_) + sum(z[r, p, t].X for r in R_ for p in P_)
    val_t = sum(v[p, k, t].X for p in P_ for k in K_)
    mat_t = sum(v[p, k, t].X for p in P_ for k in KM_)
    print(f"  Año {t}: Rec={rec_t:8,.1f} t ({rec_t/gen_t*100:4.1f}% vs meta {meta_rec[t]*100:4.1f}%) | "
          f"Val={val_t:8,.1f} t ({val_t/gen_t*100:4.1f}% vs meta {meta_val[t]*100:4.1f}%) | "
          f"Material={mat_t/val_t*100:4.1f}% (meta >=60.0%)")

# -- generar archivos de resultados en Análisis_de_Excels y resultados/ --
RES_CSV_DIR = BASE_DIR / "resultados"
RES_XLS_DIR = BASE_DIR / "Análisis_de_Excels"
RES_CSV_DIR.mkdir(exist_ok=True)
RES_XLS_DIR.mkdir(exist_ok=True)

df_costos = pd.DataFrame({
    "Componente": [
        "CP_V procesamiento variable plantas",
        "CF_P operacion fija plantas",
        "CT_3 transporte directo region-planta",
        "CT_1 transporte primario region-acopio",
        "CT_2 transporte troncal acopio-planta",
        "CF_C operacion fija centros acopio",
        "CA_S almacenaje e inventario sanitario",
        "CP_M multas deficit normativo DS8",
        "TOTAL",
    ],
    "Costo_M_CLP": [c_proc, c_fijo_p, c_tr_3, c_tr_1, c_tr_2, c_fijo_c, c_alma, c_mult, total],
    "Pct": [
        c_proc / total * 100, c_fijo_p / total * 100, c_tr_3 / total * 100,
        c_tr_1 / total * 100, c_tr_2 / total * 100, c_fijo_c / total * 100,
        c_alma / total * 100, c_mult / total * 100, 100.0,
    ],
})
df_costos.to_csv(RES_CSV_DIR / "resumen_costos.csv", index=False)

filas_flujos = []
for t in T_:
    for r in R_:
        for c in C_:
            if x[r, c, t].X > 1e-3:
                filas_flujos.append({
                    "anio": t, "tramo": "T1_Region_Acopio", "origen_id": r, "origen": nom_r[r],
                    "destino_id": c, "destino": nom_c[c], "distancia_km": dist_t1[r, c],
                    "costo_unit_M_CLP": c_t1[r, c], "toneladas": round(x[r, c, t].X, 2),
                })
    for c in C_:
        for p in P_:
            if f[c, p, t].X > 1e-3:
                filas_flujos.append({
                    "anio": t, "tramo": "T2_Acopio_Planta", "origen_id": c, "origen": nom_c[c],
                    "destino_id": p, "destino": nom_p[p], "distancia_km": dist_t2[c, p],
                    "costo_unit_M_CLP": c_t2[c, p], "toneladas": round(f[c, p, t].X, 2),
                })
    for r in R_:
        for p in P_:
            if z[r, p, t].X > 1e-3:
                filas_flujos.append({
                    "anio": t, "tramo": "T3_Region_Planta_Directo", "origen_id": r, "origen": nom_r[r],
                    "destino_id": p, "destino": nom_p[p], "distancia_km": dist_t3[r, p],
                    "costo_unit_M_CLP": c_t3[r, p], "toneladas": round(z[r, p, t].X, 2),
                })
df_flujos = pd.DataFrame(filas_flujos)
df_flujos.to_csv(RES_CSV_DIR / "flujos_transporte.csv", index=False)

filas_acopio = []
for c in C_:
    for t in T_:
        act = int(round(y[c, t].X))
        f_in = sum(x[r, c, t].X for r in R_)
        f_out = sum(f[c, p, t].X for p in P_)
        inv = s[c, t].X
        filas_acopio.append({
            "anio": t, "centro_id": c, "centro": nom_c[c], "activo": act,
            "cap_entrada_ton": cap_rec_ca[c] * act, "flujo_entrada_ton": round(f_in, 2),
            "uso_entrada_pct": round(f_in / cap_rec_ca[c] * 100, 1) if act else 0.0,
            "cap_salida_ton": cap_des_ca[c] * act, "flujo_salida_ton": round(f_out, 2),
            "uso_salida_pct": round(f_out / cap_des_ca[c] * 100, 1) if act else 0.0,
            "inv_max_ton": inv_max_ca[c] * act, "inventario_final_ton": round(inv, 2),
        })
df_acopio = pd.DataFrame(filas_acopio)
df_acopio.to_csv(RES_CSV_DIR / "capacidad_acopios.csv", index=False)

filas_plantas = []
for p in P_:
    for k in K_:
        if fact_pk[p, k] == 1:
            for t in T_:
                act = int(round(w[p, k, t].X))
                proc = v[p, k, t].X
                filas_plantas.append({
                    "anio": t, "planta_id": p, "planta": nom_p[p], "tecnologia_id": k,
                    "tecnologia": nom_k[k], "categoria_ds8": "Material" if k in KM_ else "Energetica",
                    "activa": act, "cap_efectiva_ton": cap_ef_pk[p, k] * act,
                    "valorizado_ton": round(proc, 2),
                    "uso_capacidad_pct": round(proc / cap_ef_pk[p, k] * 100, 1) if act else 0.0,
                })
df_plantas = pd.DataFrame(filas_plantas)
df_plantas.to_csv(RES_CSV_DIR / "capacidad_plantas.csv", index=False)

filas_ds8 = []
for t in T_:
    gen_t = sum(gen_nfu[r, t] for r in R_)
    rec_t = sum(x[r, c, t].X for r in R_ for c in C_) + sum(z[r, p, t].X for r in R_ for p in P_)
    val_t = sum(v[p, k, t].X for p in P_ for k in K_)
    mat_t = sum(v[p, k, t].X for p in P_ for k in KM_)
    ene_t = sum(v[p, k, t].X for p in P_ for k in KE_)
    filas_ds8.append({
        "anio": t, "generacion_ton": round(gen_t, 2),
        "meta_rec_ton": round(meta_rec[t] * gen_t, 2), "recoleccion_ton": round(rec_t, 2),
        "pct_recoleccion": round(rec_t / gen_t * 100, 2),
        "meta_val_ton": round(meta_val[t] * gen_t, 2), "valorizacion_ton": round(val_t, 2),
        "pct_valorizacion": round(val_t / gen_t * 100, 2),
        "valorizacion_material_ton": round(mat_t, 2), "valorizacion_energetica_ton": round(ene_t, 2),
        "pct_material_sobre_valorizado": round(mat_t / val_t * 100, 2),
    })
df_ds8 = pd.DataFrame(filas_ds8)
df_ds8.to_csv(RES_CSV_DIR / "cumplimiento_ds8.csv", index=False)

with pd.ExcelWriter(RES_XLS_DIR / "Resultados_Optimizacion_REP_NFU.xlsx") as writer:
    df_costos.to_excel(writer, sheet_name="Resumen_Costos_FO", index=False)
    df_ds8.to_excel(writer, sheet_name="Cumplimiento_Normativo_DS8", index=False)
    df_acopio.to_excel(writer, sheet_name="Capacidad_Centros_Acopio", index=False)
    df_plantas.to_excel(writer, sheet_name="Capacidad_Plantas_Tec", index=False)
    df_flujos.to_excel(writer, sheet_name="Matriz_Flujos_Transporte", index=False)


print(f"\nResultados exportados en {RES_CSV_DIR.name}/ y {RES_XLS_DIR.name}/")