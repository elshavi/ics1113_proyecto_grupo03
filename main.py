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
    
    # conjuntos y parametros
    J = datos["J"]
    A = datos["A"]
    D = datos["D"]
    H = datos["H"]
    cj = datos["cj"]
    beta_inicial = datos["beta"]
    
    # revisar si existe solucion
    solution_exists = modelo.Status in [gp.GRB.OPTIMAL, gp.GRB.TIME_LIMIT]

    # funcion para obtenner valor de variable
    def get_var_value(var_name):
        var = modelo.getVarByName(var_name)
        if var and solution_exists:
            try:
                # Retorna el valor de la variable
                return var.X
            except AttributeError:
                return 0.0
        return 0.0

    
    annual_data = {a: {} for a in A}
    purchase_log = [] # Almacena (año, dia, tipo, cantidad)

    # --- Recoleccion de Datos ---
    for a in A:
        # --- 1. Variables Anuales (Utilidad, Presupuesto) ---
        utilidad_a = get_var_value(f"U[{a}]")
        presupuesto_a = get_var_value(f"Pa[{a}]")

        # Inicializar sumas anuales
        total_bn_a = {j: 0.0 for j in J}
        total_b_a = {j: 0.0 for j in J}
        
        costo_bat_a = 0.0
        costo_vert_a = 0.0
        vert_total_a = 0.0
        fpr_total_a = 0.0
        fbr_total_a = 0.0
        
        # --- 2. Agregacion Diaria/Horaria (Costos, Flujos, V, BN) ---
        for d in D:
            # Costo de Baterias (Compras Diarias/Trimestrales)
            for j in J:
                var_bn_name = f"BN[{j},{a},{d}]"
                bn_x = get_var_value(var_bn_name)
                
                # Registra compra
                if bn_x > 0.001: 
                    total_bn_a[j] += bn_x
                    costo_bat_a += cj[j] * bn_x 
                    purchase_log.append((a, d, j, bn_x))
            
            # Flujos, V, Costos (Suma Horaria para el dia/año)
            for h in H:
                idx_adh = (a, d, h)
                
                V_X = get_var_value(f"V[{a},{d},{h}]")
                Fpr_X = get_var_value(f"Fpr[{a},{d},{h}]")
                
                # Fbr_X
                Fbr_X_sum_j = 0.0
                for j in J:
                    Fbr_X_sum_j += get_var_value(f"Fbr[{j},{a},{d},{h}]")
                
                vert_total_a += V_X
                fpr_total_a += Fpr_X
                fbr_total_a += Fbr_X_sum_j
                
                # Accede al costo de vertimiento (gamma)
                costo_vert_a += datos["gamma"].get(idx_adh, 0.0) * V_X
                    
        # --- 3. Capacidad Total de Bateria (B[j,a]) ---
        for j in J:
            total_b_a[j] = get_var_value(f"B[{j},{a}]")
            
        annual_data[a] = {
            "Utilidad": utilidad_a,
            "PresupuestoFinal": presupuesto_a,
            "BN": total_bn_a,
            "B": total_b_a,
            "Fpr": fpr_total_a,
            "Fbr": fbr_total_a,
            "CostoBat": costo_bat_a,
            "CostoVert": costo_vert_a,
            "Vtotal": vert_total_a
        }

    # --- Impresion del Reporte Consolidado ---
    lineas = []
    
    # 1. Encabezado
    lineas.append("=== RESUMEN OPTIMIZACION ENERGIA SOLAR Y BATERIAS ===")
    lineas.append(f"Estado del solver: {modelo.Status}")
    if modelo.ObjVal is not None:
        lineas.append(f"Valor objetivo (utilidad total): {modelo.ObjVal:.2f}")
    lineas.append(f"Horizonte: {len(A)} anos, {len(D)} dias/ano, {len(H)} horas/dia.")
    lineas.append(f"Tipos de bateria: {len(J)} (J = {J})")
    lineas.append(f"Presupuesto inicial Beta: {beta_inicial:.2f}")

    # 2. Tabla Resumen Anual Consolidado
    lineas.append("\n=== RESUMEN ANUAL CONSOLIDADO ===")

    # Crear encabezado de tabla
    header = ["Ano ", "Utilidad", "Ppto Final", "BN Total", "B Total ", "Costo Bat", "Costo Vert", "Vertida (V) ", "Fpr     ", "Fbr      "]
    lineas.append("| " + " | ".join(header) + " |")

    # Filas Anuales
    for a in A:
        data = annual_data[a]
        total_bn = sum(data["BN"].values())
        total_b = sum(data["B"].values())
        
        row = [
            f"{a:4.0f}",
            f"{data['Utilidad']:8.0f}", 
            f"{data['PresupuestoFinal']:10.0f}", 
            f"{total_bn:8.0f}",
            f"{total_b:8.0f}",
            f"{data['CostoBat']:9.0f}",
            f"{data['CostoVert']:10.2f}",
            f"{data['Vtotal']:12.2f}",
            f"{data['Fpr']:6.2f}",
            f"{data['Fbr']:6.2f}"
        ]
        lineas.append("| " + " | ".join(row) + " |")

    # 3. Detalle de Baterias por Tipo y Año
    lineas.append("\n=== DETALLE DE BATERIAS POR TIPO Y ANO ===")
    for a in A:
        total_b = sum(annual_data[a]['B'].values())
        lineas.append(f"- Ano {a:.0f} (Total Baterias en Inventario: {total_b:.0f} u.)")
        for j in J:
            bn_val = annual_data[a]['BN'][j]
            b_val = annual_data[a]['B'][j]
            lineas.append(f"  > Tipo J{j:.0f}: Compradas={bn_val:.0f} u. | Total en Inventario={b_val:.0f} u.")

    # 4. Registro Detallado de Compras (si existen)
    if purchase_log:
        lineas.append("\n=== REGISTRO DETALLADO DE COMPRAS (Dia, Tipo, Cantidad) ===")
        lineas.append("Nota: Compra total en el ano se divide en estos dias.")
        lineas.append("| Ano  | Dia  | Tipo (J) | Cantidad |")
        lineas.append("|------|------|----------|----------|")
        for a, d, j, amount in purchase_log:
            lineas.append(f"| {a:4.0f} | {d:4.0f} | {j:8.0f} | {amount:8.0f} |")
    else:
        lineas.append("\n=== REGISTRO DETALLADO DE COMPRAS: No se encontraron compras diarias. ===")

    lineas.append("=====================================================")
    print("\n".join(lineas))



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
    BN = m.addVars(J, A, D, vtype=GRB.INTEGER, lb=0.0, name="BN")

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
        #costo_baterias = gp.quicksum(cj[j] * BN[(j,a)] for j in J) este era por año, viejo
        costo_baterias = gp.quicksum(cj[j] * BN[(j,a,d)] for j in J for d in D)
        costo_vertimiento = gp.quicksum(V[(a,d,h)] * gamma[(a,d,h)] for d in D for h in H)

        m.addConstr(U[a] == ingreso - costo_baterias - costo_vertimiento, name=f"utilidad_{a}")


    # 6.2 Inventario de baterías (sumatoria de compras)
    a0 = A[0]
    for j in J:
        # caso base
        #m.addConstr(B[(j,a0)] == boj[j], name=f"base_baterias_{j}") este es por año
        m.addConstr(B[(j,a0)] == boj[j] + gp.quicksum(BN[(j,a0,d)] for d in D), name=f"base_baterias_{j}")
        # caso general
        for idx_a in range(1, len(A)):
            a_actual = A[idx_a]
            a_anterior = A[idx_a - 1]

            # este es por año
            #m.addConstr(B[(j,a_actual)] == B[(j,a_anterior)] + BN[(j,a_actual)],
                        #name=f"baterias_{j}_{a_actual}")
            
            m.addConstr(B[(j,a_actual)] == B[(j,a_anterior)] + gp.quicksum(BN[(j,a_actual,d)] for d in D),
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
        #viejo, por año
        #m.addConstr(
            #gp.quicksum(cj[j] * BN[(j,a)] for j in J) <= Pa[a],
            #name=f"limite_compra_a{a}")
        
        m.addConstr(
        gp.quicksum(cj[j] * BN[(j,a,d)] for j in J for d in D) <= Pa[a],
        name=f"limite_compra_a{a}")

    # T1: 1, T2: 91, T3: 182, T4: 274
    D_compra = [1, 91, 182, 274]

    # Restricción 6.4b: Forzar la compra de baterías (BN) a ocurrir solo en días específicos (Trimestral)
    for j in J:
        for a in A:
            for d in D:
                # D es el conjunto de días (ej. [1, 2, ..., 365]).
                # Solo permitimos que BN sea positiva si el día 'd' está en D_compra
                if d not in D_compra:
                    # Si el día no es un día de compra, BN debe ser 0.
                    m.addConstr(BN[(j,a,d)] == 0, name=f"no_compra_trimestral_{j}_{a}_{d}")
        
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
            E[(j,a0,d0,h0)] == tj[j] * boj[j] * 0.5,
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

    # La impresion de resumen_post_solve ahora contiene toda la informacion
    if modelo.Status in [gp.GRB.OPTIMAL, gp.GRB.TIME_LIMIT]:
        resumen_post_solve(modelo, datos)
    else:
        print(f"El modelo no encontro solucion factible/optima. Status={modelo.Status}")

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
         "b0": "b0",
         "SETS": "10ANOS",
         "gamma": "anualvert",
         "w": "generacion",
         "t": "t_chico",
         "c": "c_chico"
}


if __name__ == "__main__":
    datos = load_parameters(rutas, hojas)
    modelo = ejecutar_modelo(datos)