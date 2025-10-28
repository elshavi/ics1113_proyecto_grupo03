import pandas as pd 

#convertir hoja tipo (j.valor) a dict {j: valor}
def hoja_to_dict_1d(df, key_col, val_col="value"):
    diccionario = {}
    for _, row in df.iterrows():
        key = row[key_col]
        value = row[val_col]
        if pd.notnull(key):
            diccionario[key] = value
    return diccionario

#convertir hoja del tipo (a,d.valor) a dict {(a,d,h): valor}
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
    


