import gurobipy as gp
from gurobipy import GRB
import sys
from datetime import datetime
import pandas as pd 

#FUNCIONES PARA EXTRAER DATOS DEL EXCEL

#convertir hoja tipo (jvalor) a dict {j: valor}
def hoja_to_dict_1d(df, key_col, val_col="value"):
    diccionario = {}
    for _, row in df.iterrows():
        key = row[key_col]
        value = row[val_col]
        if pd.notnull(key):
            diccionario[key] = value
    return diccionario

#convertir hoja del tipo (a,d,h,valor) a dict {(a,d,h): valor}
def hoja_to_dict_3d(df, cols=("a","d","h"), val_col="value"):
    diccionario = {}
    a_col, d_col, h_col = cols
    for _, row in df.iterrows():
        a = row[a_col]
        d = row[d_col]
        h = row[h_col]
        v = row[val_col]
        if pd.notnull(a) and pd.notnull(d) and pd.notnull(h):
            diccionario[(a,d,h)] = float(v)
    return diccionario

def leer_hoja(ruta_excel, hoja):
    return pd.read_excel(ruta_excel, hoja)

def load_parameters(rutas, hojas):
    #usar hojas especificadas en hojas
    def hoja_para(parametro):
        return hojas.get(parametro, parametro)

    #Hoja sets: columnas J,A,D,H
    df_sets = leer_hoja(rutas["SETS"], hoja_para("SETS"))
    J = sorted(list(set(df_sets['J'].dropna().tolist())))
    A = sorted(list(set(df_sets['A'].dropna().tolist())))
    D = sorted(list(set(df_sets['D'].dropna().tolist())))
    H = sorted(list(set(df_sets['H'].dropna().tolist())))

    #Parametros por tipo de bateria j (1d)
    c_j_df = leer_hoja(rutas["c"], hoja_para("c"))
    etac_j_df = leer_hoja(rutas["etac"], hoja_para("etac"))
    etad_j_df = leer_hoja(rutas["etad"], hoja_para("etad"))
    t_j_df = leer_hoja(rutas["t"], hoja_para("t"))
    b0_j_df = leer_hoja(rutas["b0"], hoja_para("b0"))

    c_j = hoja_to_dict_1d(c_j_df, "j", "value")
    etac_j = hoja_to_dict_1d(etac_j_df, "j", "value")
    etad_j = hoja_to_dict_1d(etad_j_df, "j", "value")
    t_j = hoja_to_dict_1d(t_j_df, "j", "value")
    b0_j = hoja_to_dict_1d(b0_j_df, "j", "value")

    #presupuesto inicial beta
    beta_df = leer_hoja(rutas["beta"], hoja_para("beta"))
    beta_vals = beta_df['value'].dropna().tolist()
    #solo debe tener 1 valor
    if len(beta_vals) != 1:
        raise ValueError("Hoja 'beta' debe contener un único valor.")
    beta = float(beta_vals[0])

    #parametros por (a,d,h) (3d)
    p_adh_df  = leer_hoja(rutas["p"], hoja_para("p"))
    m_adh_df  = leer_hoja(rutas["m"], hoja_para("m"))
    w_adh_df  = leer_hoja(rutas["w"], hoja_para("w"))
    gamma_df = leer_hoja(rutas["gamma"], hoja_para("gamma"))

    p_adh = hoja_to_dict_3d(p_adh_df, ("a","d","h"), "value")
    m_adh = hoja_to_dict_3d(m_adh_df, ("a","d","h"), "value")
    w_adh = hoja_to_dict_3d(w_adh_df, ("a","d","h"), "value")
    gamma_adh = hoja_to_dict_3d(gamma_df, ("a","d","h"), "value")

    data = {
        "J": J,
        "A": A,
        "D": D,
        "H": H,
        "cj": c_j,
        "etacj": etac_j,
        "etadj": etad_j,
        "tj": t_j,
        "boj": b0_j,
        "beta": beta,
        "padh": p_adh,
        "madh": m_adh,
        "wadh": w_adh,
        "gamma": gamma_adh
    }

    return data

#FUNCION PARA PRINTEAR RESUMEN DE LA SOLUCIÓN OPTIMA EN CONSOLA

def resumen_post_solve(modelo, datos):
    # 1) Estado y objetivo
    status = modelo.Status
    obj = modelo.ObjVal if getattr(modelo, "SolCount", 0) > 0 else None

    # 2) Sets y parámetros clave
    J = datos["J"]
    A = datos["A"]
    D = datos["D"]
    H = datos["H"]

    beta_inicial       = datos.get("beta")
    costo_bateria      = datos.get("cj", {})
    capacidad_bateria  = datos.get("tj", {})
    baterias_iniciales = datos.get("boj", {})

    # Helper para obtener variables por prefijo (robusto a nombres largos/cortos)
    def get_var(name):
        return {v.VarName: v for v in modelo.getVars() if v.VarName.startswith(name + "[")}

    vars_B           = get_var("Baterias")          or get_var("B")
    vars_BN          = get_var("BateriasNuevas")    or get_var("BN")
    vars_Utilidad    = get_var("Utilidad")          or get_var("U")
    vars_Presupuesto = get_var("Presupuesto")       or get_var("Pa")
    vars_Vertimiento = get_var("Vertimiento")       or get_var("V")
    vars_Fpr         = get_var("Fpr")
    vars_Fbr         = get_var("Fbr")

    # 3) Agregados por año
    baterias_por_anio = {}   # {a: {j: B[j,a]}}
    nuevas_por_anio   = {}   # {a: {j: BN[j,a]}}
    utilidad_por_anio = {}   # {a: U[a]}
    presupuesto_anio  = {}   # {a: Pa[a]}
    vert_total_anio   = {}   # {a: sum_{d,h} V[a,d,h]}
    fpr_total_anio    = {}   # {a: sum_{d,h} Fpr[a,d,h]}
    fbr_total_anio    = {}   # {a: sum_{j,d,h} Fbr[j,a,d,h]}
    fbr_tipo_anio     = {}   # {a: {j: sum_{d,h} Fbr[j,a,d,h]}}

    for a in A:
        # Utilidad y presupuesto
        for key, var in vars_Utilidad.items():
            idx = key[key.find("[")+1 : key.find("]")]
            if str(a) == idx:
                utilidad_por_anio[a] = var.X

        for key, var in vars_Presupuesto.items():
            idx = key[key.find("[")+1 : key.find("]")]
            if str(a) == idx:
                presupuesto_anio[a] = var.X

        # Baterías totales y nuevas por tipo
        baterias_por_anio[a] = {}
        nuevas_por_anio[a]   = {}
        for j in J:
            for key, var in vars_B.items():
                idxs = key[key.find("[")+1 : key.find("]")].split(",")
                if len(idxs) == 2:
                    j_idx, a_idx = idxs
                    if str(j) == j_idx.strip() and str(a) == a_idx.strip():
                        baterias_por_anio[a][j] = var.X
            for key, var in vars_BN.items():
                idxs = key[key.find("[")+1 : key.find("]")].split(",")
                if len(idxs) == 2:
                    j_idx, a_idx = idxs
                    if str(j) == j_idx.strip() and str(a) == a_idx.strip():
                        nuevas_por_anio[a][j] = var.X

        # Vertimiento total año a
        total_vert_a = 0.0
        for key, var in vars_Vertimiento.items():
            idxs = key[key.find("[")+1 : key.find("]")].split(",")
            if len(idxs) == 3:
                a_idx, d_idx, h_idx = [s.strip() for s in idxs]
                if str(a) == a_idx:
                    total_vert_a += var.X
        vert_total_anio[a] = total_vert_a

        # Fpr total por año
        total_fpr_a = 0.0
        for key, var in vars_Fpr.items():
            idxs = key[key.find("[")+1 : key.find("]")].split(",")
            if len(idxs) == 3:
                a_idx, d_idx, h_idx = [s.strip() for s in idxs]
                if str(a) == a_idx:
                    total_fpr_a += var.X
        fpr_total_anio[a] = total_fpr_a

        # Fbr total por año y por tipo
        total_fbr_a = 0.0
        fbr_tipo_anio[a] = {j: 0.0 for j in J}
        for key, var in vars_Fbr.items():
            idxs = key[key.find("[")+1 : key.find("]")].split(",")
            # Formato: Fbr[j,a,d,h]
            if len(idxs) == 4:
                j_idx, a_idx, d_idx, h_idx = [s.strip() for s in idxs]
                if str(a) == a_idx:
                    total_fbr_a += var.X
                    # acumular por tipo j (permitir j no-int)
                    try:
                        j_key = int(j_idx)
                    except ValueError:
                        j_key = j_idx
                    if j_key in fbr_tipo_anio[a]:
                        fbr_tipo_anio[a][j_key] += var.X
                    else:
                        fbr_tipo_anio[a][j_key] = fbr_tipo_anio[a].get(j_key, 0.0) + var.X
        fbr_total_anio[a] = total_fbr_a

    # 4) Resumen corto en consola
    lineas = []
    lineas.append("=== RESUMEN OPTIMIZACIÓN ENERGÍA SOLAR Y BATERÍAS ===")
    lineas.append(f"Estado del solver: {status}")
    if obj is not None:
        lineas.append(f"Valor objetivo (utilidad total): {obj:.2f}")
    lineas.append(f"Horizonte: {len(A)} años, {len(D)} días/año, {len(H)} horas/día.")
    lineas.append(f"Tipos de batería: {len(J)} (J = {J})")
    lineas.append(f"Presupuesto inicial β: {beta_inicial}")

    for a in A:
        util_a = utilidad_por_anio.get(a, 0.0)
        pres_a = presupuesto_anio.get(a, 0.0)
        lineas.append(f"- Año {a}: utilidad={util_a:.2f}  presupuesto_final={pres_a:.2f}")
        # Compras nuevas por tipo
        if a in nuevas_por_anio:
            for j in J:
                if j in nuevas_por_anio[a]:
                    lineas.append(f"    Nuevas baterías tipo {j}: {nuevas_por_anio[a][j]:.2f} u.")

    # Vertimiento y flujos a red por año
    for a in A:
        lineas.append(f"- Año {a}: energía vertida = {vert_total_anio.get(a, 0.0):.4f} (∑d,h)")
        lineas.append(
            f"          energía a red: Fpr={fpr_total_anio.get(a, 0.0):.4f} (paneles→red), "
            f"Fbr={fbr_total_anio.get(a, 0.0):.4f} (baterías→red)"
        )

    lineas.append("====================================================")
    print("\n".join(lineas), flush=True)


# funcion para modelar el (a,d,h) anterior dado el (a,d,h) actual
def instante_anterior(A, D, H, a, d, h):
    h_min = H[0]
    h_max = H[-1]
    d_min = D[0]
    d_max = D[-1]
    a_min = A[0]

    # Caso 1: hora anterior en el mismo dia:
    if h != h_min:
        return (a, d, H[H.index(h) - 1])
    # si no hay hora anterior, revisar dia anterior:
    else:
        if d != d_min:
            return (a, D[D.index(d) - 1], h_max)
        # si no hay dia anterior, revisar año anterior:
        else:
            if a != a_min:
                return (A[A.index(a) - 1], d_max, h_max)
            else:
                # caso base, no hay anterior
                return None

def build_model(data: dict):

    J = data["J"]       # lista de tipos de bateria
    A = data["A"]       # lista de años
    D = data["D"]       # lista de dias
    H = data["H"]       # lista de horas

    cj    = data["cj"]      # Costo de compra e instalacion de una bateria de tipo j
    etacj = data["etacj"]   # eficiencia de carga de bateria tipo j
    etadj = data["etadj"]   # eficiencia de descarga de bateria tipo j
    tj    = data["tj"]      # maxima capacidad energetica de bateria tipo j
    boj   = data["boj"]     # cantidad de baterias j iniciales
    beta  = data["beta"]    # presupuesto inicial

    padh  = data["padh"]    # precio de venta de energia en año a, dia d, hora h
    madh  = data["madh"]    # maxima capacidad de red en año a, dia d, hora h
    wadh  = data["wadh"]    # produccion solar en año a, dia d, hora h
    gamma = data["gamma"]   # costo por energia vertida en año a, dia d, hora h


    tasa_desgaste = 0.01 / (365.0 * 24.0)  # desgaste por hora (restriccion 10)

    # 3) Crear modelo
    m = gp.Model("ENGIE_Coya_BESS")

    # 4) Variables
    # B_ja: baterías del tipo j en año a
    B = m.addVars(J, A, vtype=GRB.INTEGER, lb=0.0, name="B")

    # BN_ja: baterías nuevas del tipo j en año a
    BN = m.addVars(J, A, vtype=GRB.INTEGER, lb=0.0, name="BN")

    # Fpr_adh: flujo de paneles a red en año a, dia d y hora h
    Fpr = m.addVars(A, D, H, vtype=GRB.CONTINUOUS, lb=0.0, name="Fpr")

    # Fbr_jadh: flujo de batería j a red en año a, dia d y hora h
    Fbr = m.addVars(J, A, D, H, vtype=GRB.CONTINUOUS, lb=0.0, name="Fbr")

    # Fpb_jadh: flujo de paneles a batería j en año a, dia d y hora h
    Fpb = m.addVars(J, A, D, H, vtype=GRB.CONTINUOUS, lb=0.0, name="Fpb")

    # V_adh: energia vertida en año a, dia d y hora h
    V = m.addVars(A, D, H, vtype=GRB.CONTINUOUS, lb=0.0, name="V")

    # E_jadh: energía guardada en batería j en año a, dia d y hora h
    E = m.addVars(J, A, D, H, vtype=GRB.CONTINUOUS, lb=0.0, name="E")

    # P_a: presupuesto en año a
    Pa = m.addVars(A, vtype=GRB.CONTINUOUS, lb=0.0, name="Pa")

    # D_jadh: desgaste en baterias de tipo j en año a, dia d y hora h 
    Des = m.addVars(J, A, D, H, vtype=GRB.CONTINUOUS, lb=0.0, name="D")

    # U_a utilidad anual en año a
    U = m.addVars(A, vtype=GRB.CONTINUOUS, lb=-GRB.INFINITY, name="U")

    # 5) Función Objetivo
    # max sum_a U[a]
    m.setObjective(gp.quicksum(U[a] for a in A), GRB.MAXIMIZE)


    # 6) Restricciones
    # 6.1 Definición de utilidad anual: U[a]
    for a in A:
        # ingreso anual
        ingreso = gp.quicksum(
            padh[(a,d,h)] * (Fpr[(a,d,h)] + gp.quicksum(Fbr[(j,a,d,h)] for j in J))
            for d in D for h in H)

        # costo anual (baterías y vertimiento)
        costo_baterias = gp.quicksum(cj[j] * BN[(j,a)] for j in J)
        costo_vertimiento = gp.quicksum(V[(a,d,h)] * gamma[(a,d,h)] for d in D for h in H)

        m.addConstr(U[a] == ingreso - costo_baterias - costo_vertimiento, name=f"utilidad_{a}")


    # 6.2 Inventario de baterías (sumatoria de compras)
    a0 = A[0]
    for j in J:
        # caso base
        m.addConstr(B[(j,a0)] == boj[j], name=f"base_baterias_{j}")
        # caso general
        for idx_a in range(1, len(A)):
            a_actual = A[idx_a]
            a_anterior = A[idx_a - 1]
            m.addConstr(B[(j,a_actual)] == B[(j,a_anterior)] + BN[(j,a_actual)],
                        name=f"baterias_{j}_{a_actual}")

    # 6.3 Límite de compras por año (bmja)
    # caso base
    m.addConstr(Pa[a0] == beta + U[a0], name=f"presupuesto_base_a{a0}")
    # caso general
    for idx_a in range(1, len(A)):
        a_actual = A[idx_a]
        a_anterior = A[idx_a - 1]
        m.addConstr(Pa[a_actual] == Pa[a_anterior] + U[a_actual],
                     name=f"presupuesto_a{a_actual}")

    # 6.4 Restricción de compra de baterías por presupuesto:
    for a in A:
        m.addConstr(
            gp.quicksum(cj[j] * BN[(j,a)] for j in J) <= Pa[a],
            name=f"limite_compra_a{a}")
        
    # 6.5 Restricción de minimo flujo a red (demanda):
    for a in A:
        for d in D:
            for h in H:
                m.addConstr(
                    Fpr[(a,d,h)] +
                    gp.quicksum(Fbr[(j,a,d,h)] for j in J)
                    >= madh[(a,d,h)],
                    name=f"capacidad_red_a{a}_d{d}_h{h}")
                
    # 6.6 Restricción de energía de baterías dinámicas:
    #valores inciales
    d0 = D[0]
    h0 = H[0]

    for j in J:
        #caso base
        m.addConstr(
            E[(j,a0,d0,h0)] == 0.0,
            name=f"energia_bateria_base_{j}")
        #caso general
        for a in A:
            for d in D:
                for h in H:
                    if (a == a0) and (d == d0) and (h == h0):
                        continue
                    previo = instante_anterior(A, D, H, a, d, h)
                    a_prev, d_prev, h_prev = previo

                    m.addConstr(
                        E[(j,a,d,h)] ==
                        E[(j,a_prev,d_prev,h_prev)]
                        + etacj[j] * Fpb[(j,a,d,h)]
                        - etadj[j] * Fbr[(j,a,d,h)],
                        name=f"energia_bateria_j{j}_a{a}_d{d}_h{h}")
                    
    # 6.7 Energia maxima bateria:
    for j in J:
        for a in A:
            for d in D:
                for h in H:
                    m.addConstr(
                        E[(j,a,d,h)] <= tj[j] * B[(j,a)] - Des[(j,a,d,h)],
                        name=f"max_energia_bateria_j{j}_a{a}_d{d}_h{h}")
    
    # 6.8 La bateróa solo puede descargarse cuando tiene carga
    for j in J:
        for a in A:
            for d in D:
                for h in H:
                    m.addConstr(
                        Fbr[(j,a,d,h)] <= E[(j,a,d,h)],
                        name=f"descarga_factible_j{j}_a{a}_d{d}_h{h}")

    # 6.9 Vertimiento
    for a in A:
        for d in D:
            for h in H:
                m.addConstr(
                    V[(a,d,h)] ==
                    wadh[(a,d,h)]
                    - (Fpr[(a,d,h)] + gp.quicksum(Fpb[(j,a,d,h)] for j in J)),
                    name=f"vertimiento_a{a}_d{d}_h{h}")

    #6.10 Restricciones de desgaste
    for j in J:
        #caso base
        m.addConstr(
            Des[(j,a0,d0,h0)] == 0.0,
            name=f"desgaste_base_j{j}")
    
        #caso general
        for a in A:
            for d in D:
                for h in H:
                    if (a == a0) and (d == d0) and (h == h0):
                        continue
                    previo = instante_anterior(A, D, H, a, d, h)
                    a_prev, d_prev, h_prev = previo

                    m.addConstr(
                        Des[(j,a,d,h)] ==
                        Des[(j,a_prev,d_prev,h_prev)]
                        + tasa_desgaste * tj[j] * B[(j,a)],
                        name=f"desgaste_j{j}_a{a}_d{d}_h{h}")
                    
                    #cota superior desgaste
                    m.addConstr(
                        Des[(j,a,d,h)] <= tj[j] * B[(j,a)],
                        name=f"cota_desgaste_j{j}_a{a}_d{d}_h{h}")
                    
    return m

def ejecutar_modelo(datos: dict):

    modelo = build_model(datos)
    modelo.optimize()
    # después de modelo.optimize()

    if modelo.Status in [GRB.OPTIMAL, GRB.TIME_LIMIT]:
        resumen_post_solve(modelo, datos)
        print("\n=== Diagnóstico Económico ===")

        J = datos["J"]
        A = datos["A"]
        D = datos["D"]
        H = datos["H"]
        padh = datos["padh"]
        gamma = datos["gamma"]
        cj = datos["cj"]

        for a in A:
            ingresos = sum(
                padh[(a,d,h)] * (
                    modelo.getVarByName(f"Fpr[{a},{d},{h}]").X +
                    sum(modelo.getVarByName(f"Fbr[{j},{a},{d},{h}]").X for j in J)
                )
                for d in D for h in H
            )

            costo_bat = sum(cj[j] * modelo.getVarByName(f"BN[{j},{a}]").X for j in J)

            vert_total = sum(modelo.getVarByName(f"V[{a},{d},{h}]").X for d in D for h in H)
            costo_vert = sum(gamma[(a,d,h)] * modelo.getVarByName(f"V[{a},{d},{h}]").X for d in D for h in H)

            print(f"Año {a}: ingresos={ingresos:.2f}, costo_bat={costo_bat:.2f}, costo_vert={costo_vert:.2f}, Vtotal={vert_total:.2f}")

        print("==============================\n")
    else:
        print(f"El modelo no encontró solución factible/óptima. Status={modelo.Status}")

    return modelo
    
# Diccionario de rutas de excel para cada parámetro (independiente del SO)
rutas = {
    "SETS":  "sets.xlsx",
    "c":     "costos_baterias.xlsx",
    "etac":  "eficiencia_carga.xlsx",
    "etad":  "eficiencia_descarga.xlsx",
    "t":     "capacidad_baterias.xlsx",
    "b0":    "baterias_iniciales.xlsx",
    "beta":  "presupuesto_inicial.xlsx",
    "p":     "precio_energia.xlsx",
    "m":     "capacidad_red.xlsx",
    "w":     "produccion_solar.xlsx",
    "gamma": "costo_vertimiento.xlsx",
}

#que hoja utilizar de cada excel
#si no se pone nada, utiliza hoja con el nombre de la key
hojas = {"m": "100khogares",
         "gamma": "anualvert",
         "w": "generacion",
         "p": "preciofinal"
}


if __name__ == "__main__":
    datos = load_parameters(rutas, hojas)
    modelo = ejecutar_modelo(datos)
