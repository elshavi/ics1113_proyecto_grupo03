import gurobipy as gp
from gurobipy import GRB
from extraer_datos import load_parameters
import sys
import os
from datetime import datetime

def resumen_post_solve(modelo, datos, ruta_txt="resultado_detallado.txt"):
    """
    - Imprime en consola un resumen ejecutivo del plan óptimo.
    - Escribe un archivo .txt con más detalle de resultados.
    """

    # 1. sacar status y objetivo
    status = modelo.Status
    obj = None
    if modelo.SolCount > 0:
        obj = modelo.ObjVal

    # 2. sacar los sets y parámetros clave para reportar
    J = datos["J"]
    A = datos["A"]
    D = datos["D"]
    H = datos["H"]

    beta_inicial = datos["beta"]         # presupuesto inicial β
    costo_bateria = datos["cj"]          # cj
    capacidad_bateria = datos["tj"]      # tj
    baterias_iniciales = datos["boj"]    # b0j

    # 3. recolectar algunas decisiones relevantes del modelo
    #    - Baterias[j,a] = total de baterías tipo j instaladas en año a
    #    - BateriasNuevas[j,a] = nuevas baterías compradas en año a
    #    - Utilidad[a] = utilidad del año a
    #    - Presupuesto[a] = presupuesto en año a
    #    - Vertimiento[a,d,h] = energía vertida
    # OJO: ajustar nombres de variables si en tu modelo se llaman distinto

    # helper seguro para obtener una var por nombre base
    def get_var(name):
        return {v.VarName: v for v in modelo.getVars() if v.VarName.startswith(name + "[")}

    vars_B           = get_var("Baterias")          if any(v.VarName.startswith("Baterias[") for v in modelo.getVars()) else get_var("B")
    vars_BN          = get_var("BateriasNuevas")    if any(v.VarName.startswith("BateriasNuevas[") for v in modelo.getVars()) else get_var("BN")
    vars_Utilidad    = get_var("Utilidad")          if any(v.VarName.startswith("Utilidad[") for v in modelo.getVars()) else get_var("U")
    vars_Presupuesto = get_var("Presupuesto")       if any(v.VarName.startswith("Presupuesto[") for v in modelo.getVars()) else get_var("Pa")
    vars_Vertimiento = get_var("Vertimiento")       if any(v.VarName.startswith("Vertimiento[") for v in modelo.getVars()) else get_var("V")

    # armar data agregada por año
    baterias_por_anio = {}      # {a: {j: cantidad total instalada B[j,a]}}
    nuevas_por_anio   = {}      # {a: {j: nuevas BateriasNuevas[j,a]}}
    utilidad_por_anio = {}      # {a: Utilidad[a]}
    presupuesto_anio  = {}      # {a: Presupuesto[a]}
    vert_total_anio   = {}      # {a: vertimiento total sum_{d,h} V[a,d,h]}

    for a in A:
        # utilidad y presupuesto
        # Buscar Utilidad[a] en vars_Utilidad
        # VarName típico: Utilidad[2025] o U[2025]
        for key, var in vars_Utilidad.items():
            # extraer índice entre corchetes
            # ej: "Utilidad[2025]" -> "2025"
            idx = key[key.find("[")+1 : key.find("]")]
            if str(a) == idx:
                utilidad_por_anio[a] = var.X

        for key, var in vars_Presupuesto.items():
            idx = key[key.find("[")+1 : key.find("]")]
            if str(a) == idx:
                presupuesto_anio[a] = var.X

        # baterías totales y nuevas por tipo j
        baterias_por_anio[a] = {}
        nuevas_por_anio[a]   = {}
        for j in J:
            # Baterias[j,a]
            for key, var in vars_B.items():
                # Formato esperado: Baterias[1,2025] o B[1,2025]
                idxs = key[key.find("[")+1 : key.find("]")].split(",")
                if len(idxs) == 2:
                    j_idx, a_idx = idxs
                    if str(j) == j_idx.strip() and str(a) == a_idx.strip():
                        baterias_por_anio[a][j] = var.X

            # BateriasNuevas[j,a]
            for key, var in vars_BN.items():
                idxs = key[key.find("[")+1 : key.find("]")].split(",")
                if len(idxs) == 2:
                    j_idx, a_idx = idxs
                    if str(j) == j_idx.strip() and str(a) == a_idx.strip():
                        nuevas_por_anio[a][j] = var.X

        # vertimiento total del año a (sum_{d,h} V[a,d,h])
        total_vert_a = 0.0
        for key, var in vars_Vertimiento.items():
            # Formato: Vertimiento[2025,12,7] => a,d,h
            idxs = key[key.find("[")+1 : key.find("]")].split(",")
            if len(idxs) == 3:
                a_idx, d_idx, h_idx = idxs
                if str(a) == a_idx.strip():
                    total_vert_a += var.X
        vert_total_anio[a] = total_vert_a

    # 4. Construir texto resumen corto (para consola)
    #    Este es un párrafo humano entendible.
    resumen_lineas = []
    resumen_lineas.append("=== RESUMEN OPTIMIZACIÓN ENERGÍA SOLAR Y BATERÍAS ===")
    resumen_lineas.append(f"Estado del solver: {status}")
    if obj is not None:
        resumen_lineas.append(f"Valor objetivo (utilidad total maximizada): {obj:.2f}")
    resumen_lineas.append(f"Horizonte temporal: {len(A)} años, {len(D)} días/año, {len(H)} horas/día.")
    resumen_lineas.append(f"Tipos de batería modelados: {len(J)} (J = {J})")
    resumen_lineas.append(f"Presupuesto inicial β: {beta_inicial}")

    # año a año: utilidad, presupuesto final y nuevas compras
    for a in A:
        util_a = utilidad_por_anio.get(a, None)
        pres_a = presupuesto_anio.get(a, None)
        resumen_lineas.append(f"- Año {a}: utilidad={util_a:.2f} presupuesto_final={pres_a:.2f}")
        # compras nuevas por tipo
        if a in nuevas_por_anio:
            for j in J:
                if j in nuevas_por_anio[a]:
                    resumen_lineas.append(
                        f"    Nuevas baterías tipo {j}: {nuevas_por_anio[a][j]:.2f} unidades"
                    )

    # vertimiento total por año
    for a in A:
        resumen_lineas.append(
            f"- Año {a}: energía vertida total = {vert_total_anio[a]:.4f} (suma sobre d,h)"
        )

    resumen_lineas.append("====================================================")

    # imprimir en consola todo junto como un párrafo/ bloque
    print("\n".join(resumen_lineas), flush=True)

    # 5. Construir el reporte detallado para el .txt
    #    Metemos más contexto: definición de sets, variables y restricciones
    #    (tal como en tu formulación matemática) + resultados numéricos.

    detalle = []
    detalle.append("REPORTE COMPLETO DEL MODELO DE OPTIMIZACIÓN")
    detalle.append(f"Generado: {datetime.now().isoformat()}")
    detalle.append("")
    detalle.append("1. Estado y objetivo")
    detalle.append(f"   - Estado solver (Gurobi Status): {status}")
    if obj is not None:
        detalle.append(f"   - Objetivo total (sum_a U_a): {obj:.6f}")
    detalle.append("")
    detalle.append("2. Conjuntos")
    detalle.append(f"   J (tipos baterías): {J}")
    detalle.append(f"   A (años): {A}")
    detalle.append(f"   D (días): {D[0]}..{D[-1]} ({len(D)} días)")
    detalle.append(f"   H (horas): {H[0]}..{H[-1]} ({len(H)} horas)")
    detalle.append("")
    detalle.append("3. Parámetros clave")
    detalle.append(f"   β (presupuesto inicial): {beta_inicial}")
    detalle.append("   cj (costo compra/instalación por tipo de batería j):")
    for j in J:
        if j in costo_bateria:
            detalle.append(f"      j={j}: cj={costo_bateria[j]}")
    detalle.append("   tj (capacidad máxima de batería j):")
    for j in J:
        if j in capacidad_bateria:
            detalle.append(f"      j={j}: t_j={capacidad_bateria[j]}")
    detalle.append("   b0j (baterías iniciales por tipo j):")
    for j in J:
        if j in baterias_iniciales:
            detalle.append(f"      j={j}: b0={baterias_iniciales[j]}")
    detalle.append("")
    detalle.append("4. Resultados por año")
    for a in A:
        util_a = utilidad_por_anio.get(a, None)
        pres_a = presupuesto_anio.get(a, None)
        detalle.append(f"   Año {a}:")
        detalle.append(f"      Utilidad U_a = {util_a}")
        detalle.append(f"      Presupuesto P_a = {pres_a}")
        detalle.append(f"      Energía vertida total año {a} = {vert_total_anio[a]}")
        # baterías instaladas y nuevas
        for j in J:
            bj = baterias_por_anio.get(a, {}).get(j, None)
            bnj = nuevas_por_anio.get(a, {}).get(j, None)
            detalle.append(f"      Tipo {j}: B[j,a]={bj}  BN[j,a]={bnj}")
    detalle.append("")
    detalle.append("5. Interpretación rápida")
    detalle.append("   - La función objetivo maximiza la utilidad total anual sum_a U_a,")
    detalle.append("     donde U_a considera ingresos por venta de energía a la red,")
    detalle.append("     costos de compra de baterías y penalizaciones por vertimiento.")
    detalle.append("   - Las restricciones aseguran:")
    detalle.append("       * Balance de energía en baterías hora a hora.")
    detalle.append("       * No inyectar más potencia que la red soporta (madh).")
    detalle.append("       * Evolución del presupuesto y capacidad de compra.")
    detalle.append("       * Límite físico de almacenamiento (t_j) y desgaste (Dj).")
    detalle.append("   - El resultado indica cuántas baterías comprar cada año,")
    detalle.append("     cómo se comporta el presupuesto, y cuánta energía se vierte.")

    # 6. Guardar el reporte detallado en archivo .txt
    with open(ruta_txt, "w", encoding="utf-8") as f:
        f.write("\n".join(detalle))
        f.write("\n")

    # También puedes avisar en consola dónde quedó guardado
    print(f"\n[INFO] Reporte detallado escrito en {ruta_txt}\n", flush=True)




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
    if mip_gap != 0.00:
        modelo.Params.MIPGap = mip_gap
    if time_limit is not None:
        if time_limit != 0:
            modelo.Params.TimeLimit = time_limit

    modelo.optimize()
    # después de modelo.optimize()

    if modelo.Status in [GRB.OPTIMAL, GRB.TIME_LIMIT]:
        resumen_post_solve(modelo, datos, ruta_txt="resultado_detallado.txt")
    else:
        print(f"El modelo no encontró solución factible/óptima. Status={modelo.Status}")

    return modelo
    
# Carpeta donde está este script (src/main.py)
script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
# Carpeta del proyecto (padre de src)
project_dir = os.path.dirname(script_dir)
# Carpeta data al mismo nivel que src
data_dir = os.path.join(project_dir, "data")

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
hojas = {"m": "mx001"
}


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Uso: python main.py <mip_gap> <time_limit(seg)>")
        sys.exit(1)
    mip_gap = float(sys.argv[1])
    time_limit = int(sys.argv[2])

    datos = load_parameters(rutas, hojas)
    modelo = ejecutar_modelo(datos, mip_gap, time_limit)
    
    
