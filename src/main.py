import gurobipy as gp
from gurobipy import GRB
from extraer_datos import load_parameters
import sys
import os

# funcion para modelar el (a,d,h) anterior dado el (a,d,h) actual
def instante_anterior(A, D, H, a, d, h):
    h_min = H[0]
    h_max = H[-1]
    d_min = D[0]
    d_max = D[-1]
    a_min = A[0]

    # Caso 1: hora anterior mismo dia:
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
        ingreso = gp.quicksum(
            padh[(a,d,h)] * (Fpr[(a,d,h)] +
            gp.quicksum(Fbr[(j,a,d,h)] for j in J))
            for d in D for h in H)

        costo = gp.quicksum(
            gp.quicksum(cj[j]*BN[(j,a)] for j in J) +
            V[(a,d,h)] * gamma[(a,d,h)]
            for d in D for h in H)
        
        m.addConstr(U[a] == ingreso - costo, name=f"utilidad_{a}")

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
        
    # 6.5 Restricción de capacidad máxima de energía de red:
    for a in A:
        for d in D:
            for h in H:
                m.addConstr(
                    Fpr[(a,d,h)] +
                    gp.quicksum(Fbr[(j,a,d,h)] for j in J)
                    <= madh[(a,d,h)],
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

def ejecutar_modelo(datos: dict, mip_gap = 0.0, time_limit = None):

    modelo = build_model(datos)

    # gurobi config
    modelo.Params.MIPGap = mip_gap
    if time_limit is not None:
        modelo.Params.TimeLimit = time_limit

    modelo.optimize()
    return modelo
    
# Carpeta donde está este script (main.py)
# sys.argv[0] es la ruta del script que se está ejecutando
base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))

# Carpeta 'data' al lado del script
data_dir = os.path.join(base_dir, "data")

# Diccionario de rutas de excel para cada parámetro (independiente del SO)
rutas = {
    "SETS":  os.path.join(data_dir, "sets.xlsx"),
    "c":     os.path.join(data_dir, "costos_baterias.xlsx"),
    "etac":  os.path.join(data_dir, "eficiencia_carga.xlsx"),
    "etad":  os.path.join(data_dir, "eficiencia_descarga.xlsx"),
    "t":     os.path.join(data_dir, "capacidad_baterias.xlsx"),
    "b0":    os.path.join(data_dir, "baterias_iniciales.xlsx"),
    "beta":  os.path.join(data_dir, "presupuesto_inicial.xlsx"),
    "p":     os.path.join(data_dir, "precio_energia.xlsx"),
    "m":     os.path.join(data_dir, "capacidad_red.xlsx"),
    "w":     os.path.join(data_dir, "produccion_solar.xlsx"),
    "gamma": os.path.join(data_dir, "costo_vertimiento.xlsx"),
}

#que hoja utilizar de cada excel
#si no se pone nada, utiliza hoja con el nombre de la key
hojas = {
}


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Uso: python main.py <mip_gap> <time_limit(seg)>")
        sys.exit(1)
    mip_gap = float(sys.argv[1])
    time_limit = int(sys.argv[2])

    datos = load_parameters(rutas, hojas)
    modelo = ejecutar_modelo(datos, mip_gap, time_limit)
    
    
