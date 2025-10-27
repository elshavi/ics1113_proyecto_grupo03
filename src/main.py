import gurobipy as gp
from gurobipy import GRB
from io_params import load_all_params

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

    padh  = data["padh"]    
    madh  = data["madh"]
    wadh  = data["wadh"]
    gamma = data["gamma"]

    tasa_desgaste = 0.01 / (365.0 * 24.0)  # desgaste por hora (restriccion 10)

    # 3) Crear modelo
    m = gp.Model("ENGIE_Coya_BESS")

    # 4) Variables
    # B_ja: baterías del tipo j en año a
    B = m.addVars(J, A, vtype=GRB.INTEGER, lb=0, name="B")

    # BN_ja: baterías nuevas del tipo j en año a
    BN = m.addVars(J, A, vtype=GRB.INTEGER, lb=0, name="BN")

    # Fpr_adh: flujo de paneles a red en año a, dia d y hora h
    Fpr = m.addVars(A, D, H, lb=0.0, name="Fpr")

    # Fbr_jadh: flujo de batería j a red en año a, dia d y hora h
    Fbr = m.addVars(J, A, D, H, lb=0.0, name="Fbr")

    # Fpb_jadh: flujo de paneles a batería j en año a, dia d y hora h
    Fpb = m.addVars(J, A, D, H, lb=0.0, name="Fpb")

    # V_adh: energia vertida en año a, dia d y hora h
    V = m.addVars(A, D, H, lb=0.0, name="V")

    # E_jadh: energía guardada en batería j en año a, dia d y hora h
    E = m.addVars(J, A, D, H, lb=0.0, name="E")

    # P_a: presupuesto en año a
    Pa = m.addVars(A, lb=0.0, name="Pa")

    # D_jadh: desgaste en baterias de tipo j en año a, dia d y hora h 
    D = m.addVars(J, A, D, H, lb=0.0, name="D")

    # 5) Función Objetivo
    # max sum_{a,d,h} padh*(Fpr + sum_j Fbr) - sum_{j,a} cj*BN(j,a) - sum_{a,d,h} gamma*V
    rev = gp.quicksum(padh[(a,d,h)] * (Fpr[(a,d,h)] +
                gp.quicksum(Fbr[(j,a,d,h)] for j in J))
                for a in A for d in D for h in H)

    cost_buy = gp.quicksum(cj[j] * BN[(j,a)] for j in J for a in A)
    pen_vert = gp.quicksum(gamma[(a,d,h)] * V[(a,d,h)] for a in A for d in D for h in H)

    m.setObjective(rev - cost_buy - pen_vert, GRB.MAXIMIZE)

    # 6) Restricciones

    # 6.1 Inventario de baterías (sumatoria de compras)
    for j in J:
        # Caso base: Bj,a0 = boj[j] en el primer año
        first_a = A[0]
        m.addConstr(B[(j, first_a)] == boj[j] + BN[(j, first_a)],
                    name=f"inv_base_{j}")

        for a_prev, a_cur in zip(A[:-1], A[1:]):
            m.addConstr(B[(j, a_cur)] == B[(j, a_prev)] + BN[(j, a_cur)],
                        name=f"inv_{j}_{a_cur}")

    # 6.2 Límite de compras por año (bmja)
    for j in J:
        for a in A:
            if (j, a) in bmja:
                m.addConstr(BN[(j,a)] <= bmja[(j,a)], name=f"bmja_{j}_{a}")

    # 6.3 Límite de capacidad de red: Fpr + sum_j Fbr <= madh
    for a in A:
        for d in D:
            for h in H:
                m.addConstr(Fpr[(a,d,h)] +
                            gp.quicksum(Fbr[(j,a,d,h)] for j in J)
                            <= madh[(a,d,h)], name=f"cap_red_{a}_{d}_{h}")

    # 6.4 Dinámica de energía en baterías:
    # E(j,a,d,h) = E(j,a,d,h-1) + etacj[j]*Fpb(j,a,d,h) - etadj[j]*Fbr(j,a,d,h)
    # Con “carry” entre horas/días/años (aquí lo encadenamos h->h+1, d->d+1, a->a+1)
    def prev_time(a, d, h):
        # regresa (a',d',h') anterior (h-1), si h=0 retrocede día; si d=1 retrocede año
        if h > 0:
            return (a, d, h-1)
        else:
            if d > 1:
                return (a, d-1, 23)
            else:
                # paso de año: del 1/0 al año anterior día 365 hora 23
                # (asumimos 365 días)
                idx = A.index(a)
                if idx == 0:
                    return None  # el primer tiempo no tiene anterior
                a_prev = A[idx-1]
                return (a_prev, 365, 23)

    for j in J:
        for a in A:
            for d in D:
                for h in H:
                    prev = prev_time(a,d,h)
                    if prev is None:
                        # Caso base: E=0 al inicio
                        m.addConstr(E[(j,a,d,h)] ==
                                    0.0 + etacj[j]*Fpb[(j,a,d,h)] - etadj[j]*Fbr[(j,a,d,h)],
                                    name=f"E_base_{j}_{a}_{d}_{h}")
                    else:
                        ap, dp, hp = prev
                        m.addConstr(E[(j,a,d,h)] ==
                                    E[(j,ap,dp,hp)] + etacj[j]*Fpb[(j,a,d,h)] - etadj[j]*Fbr[(j,a,d,h)],
                                    name=f"E_dyn_{j}_{a}_{d}_{h}")

    # 6.5 Energía máxima por tipo de batería: E(j,.) <= tj[j]*B(j,a)
    for j in J:
        for a in A:
            for d in D:
                for h in H:
                    m.addConstr(E[(j,a,d,h)] <= tj[j] * B[(j,a)],
                                name=f"E_cap_{j}_{a}_{d}_{h}")

    # 6.6 Solo descargar si hay energía: Fbr <= E
    for j in J:
        for a in A:
            for d in D:
                for h in H:
                    m.addConstr(Fbr[(j,a,d,h)] <= E[(j,a,d,h)],
                                name=f"discharge_bound_{j}_{a}_{d}_{h}")

    # 6.7 Definición de vertimiento: V = wadh - (Fpr + sum_j Fpb)
    for a in A:
        for d in D:
            for h in H:
                m.addConstr(
                    V[(a,d,h)] == wadh[(a,d,h)] - (
                        Fpr[(a,d,h)] + gp.quicksum(Fpb[(j,a,d,h)] for j in J)
                    ),
                    name=f"vert_{a}_{d}_{h}"
                )

    # 6.8 (Opcional) Presupuesto anual si lo quieren explícito:
    # Pyear[a0] = beta - sum_j cj*BN(j,a0) + sum_{d,h} padh*(Fpr + sum_j Fbr) - sum_{d,h} gamma*V
    # Pyear[a]  = Pyear[a-1] - compras(a) + ingresos(a) - penalizaciones(a)
    a0 = A[0]
    ingresos_a0 = gp.quicksum(padh[(a0,d,h)] * (Fpr[(a0,d,h)] +
                    gp.quicksum(Fbr[(j,a0,d,h)] for j in J)) for d in D for h in H)
    compras_a0 = gp.quicksum(cj[j]*BN[(j,a0)] for j in J)
    pen_a0 = gp.quicksum(gamma[(a0,d,h)]*V[(a0,d,h)] for d in D for h in H)
    m.addConstr(Pyear[a0] == beta - compras_a0 + ingresos_a0 - pen_a0, name="budget_base")

    for a_prev, a_cur in zip(A[:-1], A[1:]):
        ingresos = gp.quicksum(padh[(a_cur,d,h)] * (Fpr[(a_cur,d,h)] +
                    gp.quicksum(Fbr[(j,a_cur,d,h)] for j in J)) for d in D for h in H)
        compras  = gp.quicksum(cj[j]*BN[(j,a_cur)] for j in J)
        pen      = gp.quicksum(gamma[(a_cur,d,h)]*V[(a_cur,d,h)] for d in D for h in H)
        m.addConstr(Pyear[a_cur] == Pyear[a_prev] - compras + ingresos - pen,
                    name=f"budget_{a_cur}")

    # (Si desean limitar compras por presupuesto disponible en el año:)
    # BN_cost(a) <= Pyear(a-1) + ingresos(a)  -> pueden ajustar según su lógica
    return m

if __name__ == "__main__":
    # Carpeta con los xlsx (uno por parámetro)
    base_params_dir = "./params_excel"
    model = build_model(base_params_dir)
    model.optimize()

    if model.status == GRB.OPTIMAL or model.status == GRB.INTERRUPTED:
        print("FO =", model.objVal)
