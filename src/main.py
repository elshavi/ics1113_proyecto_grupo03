import gurobipy as gp
from gurobipy import GRB
from io_params import load_all_params

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
    D = m.addVars(J, A, D, H, vtype=GRB.CONTINUOUS, lb=0.0, name="D")

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
    m.addConstrs(Pa[a0] == beta + U[a0], name=f"presupuesto_base_a{a0}")
    # caso general
    for idx_a in range(1, len(A)):
        a_actual = A[idx_a]
        a_anterior = A[idx_a - 1]
        m.addConstrs(Pa[a_actual] == Pa[a_anterior] + U[a_actual],
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
                    
    # 6.7 Restricción de capacidad máxima de baterías considerando desgaste:

if __name__ == "__main__":
    # Carpeta con los xlsx (uno por parámetro)
    base_params_dir = "./params_excel"
    model = build_model(base_params_dir)
    model.optimize()

    if model.status == GRB.OPTIMAL or model.status == GRB.INTERRUPTED:
        print("FO =", model.objVal)
