import streamlit as st
import pandas as pd
import datetime
import sqlite3
import os
import io
import openpyxl

# 1. CONFIGURAÇÃO INICIAL
st.set_page_config(page_title="Pesagem e Balanceamento", layout="wide", initial_sidebar_state="expanded")

if 'usuario_logado' not in st.session_state:
    st.session_state['usuario_logado'] = False
    st.session_state['nivel_acesso'] = 0 
    st.session_state['nome_usuario'] = ""
if 'pagina_atual' not in st.session_state:
    st.session_state['pagina_atual'] = 'consulta'

def safe_float(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        s = str(val).strip()
        if not s: return 0.0
        if '.' in s and ',' in s: s = s.replace('.', '').replace(',', '.')
        elif ',' in s: s = s.replace(',', '.')
        return float(s)
    except:
        return 0.0

def safe_str(val):
    return "" if pd.isna(val) else str(val).strip()

# 2. CONEXÃO E CARREGAMENTO DOS BANCOS DE DADOS
@st.cache_data
def carregar_dados_banco():
    db_path = 'aeronaves.db'
    excel_path = 'Cópia de Ficha_Pesagem_v2.xlsm'
    
    if not os.path.exists(db_path) and os.path.exists(excel_path):
        try:
            df_excel = pd.read_excel(excel_path, sheet_name='Planilha1')
            conn = sqlite3.connect(db_path)
            df_excel.to_sql('pesagens', conn, if_exists='replace', index=False)
            conn.close()
        except Exception as e:
            st.error(f"Erro ao converter a planilha para o banco de dados: {e}")

    if os.path.exists(db_path):
        try:
            conn = sqlite3.connect(db_path)
            df = pd.read_sql("SELECT * FROM pesagens", conn)
            conn.close()
            
            if 'revisao' in df.columns and 'Revisao' not in df.columns:
                df = df.rename(columns={'revisao': 'Revisao'})
            if 'Pesagem' in df.columns:
                df['Pesagem'] = df['Pesagem'].astype(str)
            if 'Revisao' in df.columns:
                df['Revisao'] = df['Revisao'].astype(str)
            return df
        except Exception as e:
            st.error(f"Erro ao ler o banco de dados SQLite: {e}")
            
    return pd.DataFrame()

@st.cache_data
def carregar_tipos_aeronave():
    excel_path = 'Cópia de Ficha_Pesagem_v2.xlsm'
    if not os.path.exists(excel_path):
        st.warning(f"Planilha não encontrada: '{excel_path}'.")
        return {}
    try:
        df_tipos = pd.read_excel(excel_path, sheet_name='Sheet1')
        
        col_modelo = 'MODEL' if 'MODEL' in df_tipos.columns else df_tipos.columns[0]
        col_prefixo = 'VRG REG.' if 'VRG REG.' in df_tipos.columns else df_tipos.columns[1]
        col_armnose = 'ARMNOSE' if 'ARMNOSE' in df_tipos.columns else None
        
        col_vrbl = next((c for c in df_tipos.columns if 'VRBL' in str(c).upper() or 'VRBL NUMBER' in str(c).upper()), None)
        col_serial = next((c for c in df_tipos.columns if 'SERIAL' in str(c).upper()), None)
        col_line = next((c for c in df_tipos.columns if 'LINE' in str(c).upper()), None)
        
        dict_aero = {}
        for _, row in df_tipos.iterrows():
            prefixo = str(row[col_prefixo]).strip()
            if prefixo and prefixo != 'nan':
                modelo = str(row[col_modelo]).strip()
                arm_nose = 93.0000
                if col_armnose and pd.notna(row[col_armnose]):
                    try: arm_nose = float(row[col_armnose])
                    except: pass
                
                dict_aero[prefixo] = {
                    'modelo': modelo, 
                    'armnose': arm_nose,
                    'vrbl': str(row[col_vrbl]) if col_vrbl and pd.notna(row[col_vrbl]) else "",
                    'serial': str(row[col_serial]) if col_serial and pd.notna(row[col_serial]) else "",
                    'line': str(row[col_line]) if col_line and pd.notna(row[col_line]) else ""
                }
        return dict_aero
    except Exception as e:
        return {}

df_historico = carregar_dados_banco()
dict_tipos_aeronave = carregar_tipos_aeronave()

# 3. FUNÇÃO DE EXPORTAÇÃO PARA EXCEL
def gerar_excel_por_template(dados, caminho_template="exemplo_ficha.xlsx"):
    try:
        wb = openpyxl.load_workbook(caminho_template)
    except FileNotFoundError:
        st.error(f"O arquivo de template '{caminho_template}' não foi encontrado na pasta.")
        return b""

    ws = wb['Ficha de pesagem - Template']
    
    # CABEÇALHO
    ws['G8'] = dados.get('prefixo', '')
    ws['V8'] = dados.get('pesado_por', '')
    ws['G9'] = dados.get('modelo', '')
    ws['V9'] = dados.get('local', '')
    ws['V10'] = dados.get('data', '')
    ws['V12'] = dados.get('razao', '')
    ws['G13'] = dados.get('config_lopa', '')
    ws['G14'] = dados.get('lopa', '')
    ws['S22'] = dados.get('arm_a', '')
    ws['S24'] = dados.get('arm_b_lh', '')
    ws['S25'] = dados.get('arm_b_rh', '')
    
    ws['P43'] = dados.get('issue_date', '') 
    ws['O65'] = dados.get('ultima_pesagem', '') 
    ws['W65'] = dados.get('ultima_revisao', '') 
    
    ws['G10'] = dados.get('vrbl_number', '') 
    ws['G11'] = dados.get('serial_number', '') 
    ws['G12'] = dados.get('line_number', '') 
    
    ws['X43'] = dados.get('numero_pesagem', '') 
    ws['AC43'] = dados.get('revisao', '') 
    
    # REACTIONS & BASIC WEIGHT
    reacoes = ['LH', 'RH', 'NOSE', 'TAIL', 'TOTAL REGISTERED', 'DEDUCTIONS', 'ADDITIONS', 'AIRCRAFT BASIC WEIGHT']
    linha_reacao = 31
    for r in reacoes:
        rd = dados['reactions'].get(r, {})
        ws[f'N{linha_reacao}'] = rd.get('weight', 0.0)
        ws[f'S{linha_reacao}'] = rd.get('arm', 0.0)
        ws[f'X{linha_reacao}'] = rd.get('moment', 0.0)
        linha_reacao += 1

    ws['J44'] = dados.get('cg_mac', 0.0)

    def set_cell_value(sheet, row, col_letter, value):
        cell = sheet[f"{col_letter}{row}"]
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            for range_merged in sheet.merged_cells.ranges:
                if f"{col_letter}{row}" in range_merged:
                    min_col, min_row, _, _ = range_merged.bounds
                    cell = sheet.cell(row=min_row, column=min_col)
                    break
        cell.value = value

    # DEDUCTIONS
    deducoes = dados.get('deductions', [])
    for idx in range(70, 94):
        if (idx - 70) < len(deducoes):
            d = deducoes[idx - 70]
            desc = d.get('desc', '')
            w_val = d.get('w', 0.0)
            a_val = d.get('a', 0.0)
            m_val = d.get('m', 0.0)
        else:
            desc, w_val, a_val, m_val = "", 0.0, 0.0, 0.0
            
        set_cell_value(ws, idx, 'B', desc)
        set_cell_value(ws, idx, 'R', w_val)
        set_cell_value(ws, idx, 'V', a_val)
        set_cell_value(ws, idx, 'Z', m_val)

    # ADDITIONS
    adicoes = dados.get('additions', [])
    for idx in range(96, 127):
        if (idx - 96) < len(adicoes):
            a = adicoes[idx - 96]
            desc = a.get('desc', '')
            w_val = a.get('w', 0.0)
            a_val = a.get('a', 0.0)
            m_val = a.get('m', 0.0)
        else:
            desc, w_val, a_val, m_val = "", 0.0, 0.0, 0.0
            
        set_cell_value(ws, idx, 'B', desc)
        set_cell_value(ws, idx, 'R', w_val)
        set_cell_value(ws, idx, 'V', a_val)
        set_cell_value(ws, idx, 'Z', m_val)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# 4. TELAS E FUNCIONALIDADES

# Adicionada de volta a função da tela de consulta que havia sido apagada
def tela_consulta():
    st.title("Consulta de Histórico")
    prefixos_unicos = sorted(df_historico['Prefixo'].dropna().unique().tolist()) if 'Prefixo' in df_historico.columns else []
    prefixo = st.selectbox("Aeronave", [""] + prefixos_unicos)
    
    if prefixo:
        df_filtrado = df_historico[df_historico['Prefixo'] == prefixo]
        pesagens = sorted(df_filtrado['Pesagem'].dropna().unique().tolist())
        pesagem = st.selectbox("Pesagem", [""] + pesagens)
        
        if pesagem:
            revisoes = sorted(df_filtrado[df_filtrado['Pesagem'] == pesagem]['Revisao'].dropna().unique().tolist())
            revisao = st.selectbox("Revisão", [""] + revisoes)
            
            if revisao:
                st.divider()
                linha = df_filtrado[(df_filtrado['Pesagem'] == pesagem) & (df_filtrado['Revisao'] == revisao)].iloc[0]
                renderizar_ficha_visualizacao(prefixo, pesagem, revisao, linha)

def renderizar_ficha_visualizacao(prefixo, pesagem, revisao, row):
    st.subheader(f"Ficha Técnica: {prefixo} (Pesagem: {pesagem} | Revisão: {revisao})")
    aba1, aba2, aba3, aba4, aba5 = st.tabs(["Dados Gerais", "Células de Carga", "Deductions", "Additions", "Weighing Report"])
    info_aero = dict_tipos_aeronave.get(prefixo, {})

    with aba1:
        c_p, c_r = st.columns(2)
        c_p.text_input("Número da Pesagem:", value=safe_str(row.get('Pesagem', '')), disabled=True)
        c_r.text_input("Número da Revisão:", value=safe_str(row.get('Revisao', '')), disabled=True)

        c1, c2, c3 = st.columns(3)
        c1.text_input("Data de emissão:", value=safe_str(row.get('Data da ficha', row.get('Data_da_Pesagem', ''))), disabled=True)
        c2.text_input("Pesado por:", value=safe_str(row.get('Pesado Por', row.get('WEIGHED BY', ''))), disabled=True)
        c3.text_input("Local da pesagem:", value=safe_str(row.get('Local da pesagem', '')), disabled=True)
        
        c4, c5, c6 = st.columns(3)
        c4.text_input("Data da pesagem:", value=safe_str(row.get('Data_da_Pesagem', '')), disabled=True)
        c5.text_input("LOPA:", value=safe_str(row.get(' LOPA', row.get('LOPA', ''))), disabled=True)
        c6.text_input("Configuração LOPA:", value=safe_str(row.get('Configuração LOPA ', '')), disabled=True)
        
        c7, c8, c9 = st.columns(3)
        c7.text_input("VRBL. NUMBER:", value=safe_str(row.get('VRBL', info_aero.get('vrbl', ''))), disabled=True)
        c8.text_input("SERIAL NUMBER:", value=safe_str(row.get('SERIAL', info_aero.get('serial', ''))), disabled=True)
        c9.text_input("LINE NUMBER:", value=safe_str(row.get('LINE', info_aero.get('line', ''))), disabled=True)
        st.text_input("Razão para emissão:", value=safe_str(row.get('Motivo', '')), disabled=True)

    with aba2:
        p1_c1, p1_c2, p1_c3 = st.columns(3)
        with p1_c1:
            st.text_input("Nariz LH", value=safe_str(row.get('Peso nariz LH', '')), disabled=True)
            st.text_input("Nariz RH", value=safe_str(row.get('Peso Nariz RH', '')), disabled=True)
        with p1_c2:
            st.text_input("RH MLG", value=safe_str(row.get('Peso MLG RH 1 ', row.get('Peso MLG RH 1', ''))), disabled=True)
        with p1_c3:
            st.text_input("LH MLG", value=safe_str(row.get('Peso MLG LH 1', '')), disabled=True)
            
        st.write("---")
        p2_c1, p2_c2, p2_c3 = st.columns(3)
        with p2_c1:
            st.text_input("Nariz LH P2", value=safe_str(row.get('Peso nariz LH pesagem 2', '')), disabled=True)
            st.text_input("Nariz RH P2", value=safe_str(row.get('Peso Nariz RH pesagem 2', '')), disabled=True)
        with p2_c2:
            st.text_input("RH MLG P2", value=safe_str(row.get('Peso MLG RH 2 ', row.get('Peso MLG RH 2', ''))), disabled=True)
        with p2_c3:
            st.text_input("LH MLG P2", value=safe_str(row.get('Peso MLG LH2 pesagem 2', '')), disabled=True)

    with aba3:
        deducoes_lista = []
        for i in range(1, 16):
            desc = row.get(f'Deductions description {i}', None)
            peso = row.get(f'Deductions Weigth {i}', None)
            arm = row.get(f'Deductions arm {i}', None)
            if pd.notna(desc) and str(desc).strip() != '':
                deducoes_lista.append({"Descrição": desc, "Peso [Kg]": safe_float(peso), "Arm [pol]": safe_float(arm)})
        if deducoes_lista:
            st.dataframe(pd.DataFrame(deducoes_lista), use_container_width=True, hide_index=True)

    with aba4:
        adicoes_lista = []
        for i in range(1, 17):
            desc = row.get(f'Additions Description {i}', None)
            peso = row.get(f'Additions weigth {i}', None)
            arm = row.get(f'Additions arm {i}', None)
            if pd.notna(desc) and str(desc).strip() != '':
                adicoes_lista.append({"Descrição": desc, "Peso [Kg]": safe_float(peso), "Arm [pol]": safe_float(arm)})
        if adicoes_lista:
            st.dataframe(pd.DataFrame(adicoes_lista), use_container_width=True, hide_index=True)

    with aba5:
        try:
            lh_1 = safe_float(row.get('Peso MLG LH 1', 0))
            lh_2 = safe_float(row.get('Peso MLG LH2', 0))
            lh_1_p2 = safe_float(row.get('Peso MLG LH 1 pesagem 2', 0))
            lh_2_p2 = safe_float(row.get('Peso MLG LH2 pesagem 2', 0))
            rh_1 = safe_float(row.get('Peso MLG RH 1 ', 0))
            rh_2 = safe_float(row.get('Peso MLG RH 2', 0))
            rh_1_p2 = safe_float(row.get('Peso MLG RH 1  pesagem 2', 0))
            rh_2_p2 = safe_float(row.get('Peso MLG RH 2 pesagem 2', 0))
            nose_lh_1 = safe_float(row.get('Peso nariz LH', 0))
            nose_lh_2 = safe_float(row.get('Peso nariz LH pesagem 2', 0))
            nose_rh_1 = safe_float(row.get('Peso Nariz RH', 0))
            nose_rh_2 = safe_float(row.get('Peso Nariz RH pesagem 2', 0))
            
            nariz_lh = (nose_lh_1 + nose_lh_2)/2
            nariz_rh = (nose_rh_1 + nose_rh_2)/2
            nose = nariz_lh + nariz_rh
            
            tail = safe_float(row.get('Tail', 0))
            canela_mlg_lh = safe_float(row.get('Canela MLG LH', 0))
            canela_mlg_rh = safe_float(row.get('Canela MLG RH', 0))
        except:
            lh, rh, nose, tail = 0.0, 0.0, 0.0, 0.0
            canela_mlg_lh, canela_mlg_rh = 0.0, 0.0

        arm_b_lh = (canela_mlg_lh * 0.1304) + 706.17
        arm_b_rh = (canela_mlg_rh * 0.1304) + 706.17
        tipo_aeronave = info_aero.get('modelo', "")
        arm_a = info_aero.get('armnose', 93.0000)
        arm_c = 1298.9000

        rh = (rh_1 + rh_2 + rh_1_p2 + rh_2_p2) / 2
        lh = (lh_1 + lh_2 + lh_1_p2 + lh_2_p2) / 2

        m_lh = lh * arm_b_lh
        m_rh = rh * arm_b_rh
        m_nose = nose * arm_a
        m_tail = tail * arm_c

        tot_reg_weight = lh + rh + nose + tail
        tot_reg_moment = m_lh + m_rh + m_nose + m_tail
        tot_reg_arm = tot_reg_moment / tot_reg_weight if tot_reg_weight > 0 else 0.0

        ded_weight = sum([item["Peso [Kg]"] for item in deducoes_lista])
        ded_moment = sum([item["Peso [Kg]"] * item["Arm [pol]"] for item in deducoes_lista])
        ded_arm = ded_moment / ded_weight if ded_weight > 0 else 0.0

        momento_extra_flaps = 0
        flaps_ativo = any("Flaps" in str(item.get("Descrição", "")) for item in adicoes_lista)
        if flaps_ativo:
            if tipo_aeronave == "B737-MAX": momento_extra_flaps = 6190
            elif tipo_aeronave in ["B737-800", "B737-800SFP", "737-800", "B737-700"]: momento_extra_flaps = 5930

        add_weight = sum([item["Peso [Kg]"] for item in adicoes_lista])
        add_moment = sum([item["Peso [Kg]"] * item["Arm [pol]"] for item in adicoes_lista]) + momento_extra_flaps
        add_arm = add_moment / add_weight if add_weight > 0 else 0.0

        basic_weight = tot_reg_weight + add_weight - ded_weight
        basic_moment = tot_reg_moment + add_moment - ded_moment
        basic_arm = basic_moment / basic_weight if basic_weight > 0 else 0.0
        cg_mac = ((basic_arm - 627.1) / 1.558) if basic_arm > 0 else 0.0

        report_data = {
            "Reaction / Item": ["LH", "RH", "NOSE", "TAIL", "TOTAL REGISTERED", "DEDUCTIONS", "ADDITIONS", "AIRCRAFT BASIC WEIGHT"],
            "Weight (kg)": [lh, rh, nose, tail, tot_reg_weight, ded_weight, add_weight, basic_weight],
            "Arm (inch)": [arm_b_lh, arm_b_rh, arm_a, arm_c, tot_reg_arm, ded_arm, add_arm, basic_arm],
            "Moment (kg x inch)": [m_lh, m_rh, m_nose, m_tail, tot_reg_moment, ded_moment, add_moment, basic_moment]
        }
        st.dataframe(pd.DataFrame(report_data), use_container_width=True, hide_index=True)

        st.divider()
        col_res1, col_res2 = st.columns(2)
        with col_res1:
            st.metric(label="Aircraft Basic Weight", value=f"{basic_weight:,.4f} kg")
            st.metric(label="Aircraft Basic Arm", value=f"{basic_arm:,.4f} in")
        with col_res2:
            st.metric(label="C.G. (% MAC)", value=f"{cg_mac:.2f} %")
            
        pesagem_atual = safe_float(row.get('Pesagem', 0))
        revisao_atual = safe_float(row.get('Revisao', 0))
        
        df_aero = df_historico[df_historico['Prefixo'] == prefixo].copy()
        df_aero['Pesagem_num'] = pd.to_numeric(df_aero['Pesagem'], errors='coerce').fillna(0)
        df_aero['Revisao_num'] = pd.to_numeric(df_aero['Revisao'], errors='coerce').fillna(0)
        df_prev = df_aero[(df_aero['Pesagem_num'] < pesagem_atual) | ((df_aero['Pesagem_num'] == pesagem_atual) & (df_aero['Revisao_num'] < revisao_atual))]
        df_prev = df_prev.sort_values(by=['Pesagem_num', 'Revisao_num'])
        
        ultima_p = str(int(df_prev.iloc[-1]['Pesagem_num'])) if not df_prev.empty else ""
        ultima_r = str(int(df_prev.iloc[-1]['Revisao_num'])) if not df_prev.empty else ""

        dados_excel = {
            "prefixo": prefixo, "modelo": tipo_aeronave, "pesado_por": safe_str(row.get('Pesado Por', '')),
            "local": safe_str(row.get('Local da pesagem', '')), "data": safe_str(row.get('Data_da_Pesagem', '')),
            "config_lopa": safe_str(row.get('Configuração LOPA ', '')), "lopa": safe_str(row.get(' LOPA', '')),
            "razao": safe_str(row.get('Motivo', '')), "cg_mac": cg_mac,
            "arm_a": arm_a, "arm_b_lh": arm_b_lh, "arm_b_rh": arm_b_rh,
            "issue_date": safe_str(row.get('Data da ficha', '')),
            "ultima_pesagem": ultima_p, "ultima_revisao": ultima_r, 
            "vrbl_number": safe_str(row.get('VRBL', info_aero.get('vrbl', ''))),
            "serial_number": safe_str(row.get('SERIAL', info_aero.get('serial', ''))),
            "line_number": safe_str(row.get('LINE', info_aero.get('line', ''))),
            "numero_pesagem": str(int(pesagem_atual)) if pesagem_atual else "",
            "revisao": str(int(revisao_atual)) if revisao_atual else "",
            "reactions": {
                'LH': {'weight': lh, 'arm': arm_b_lh, 'moment': m_lh},
                'RH': {'weight': rh, 'arm': arm_b_rh, 'moment': m_rh},
                'NOSE': {'weight': nose, 'arm': arm_a, 'moment': m_nose},
                'TAIL': {'weight': tail, 'arm': arm_c, 'moment': m_tail},
                'TOTAL REGISTERED': {'weight': tot_reg_weight, 'arm': tot_reg_arm, 'moment': tot_reg_moment},
                'DEDUCTIONS': {'weight': ded_weight, 'arm': ded_arm, 'moment': ded_moment},
                'ADDITIONS': {'weight': add_weight, 'arm': add_arm, 'moment': add_moment},
                'AIRCRAFT BASIC WEIGHT': {'weight': basic_weight, 'arm': basic_arm, 'moment': basic_moment}
            },
            "deductions": [{'desc': d["Descrição"], 'w': d["Peso [Kg]"], 'a': d["Arm [pol]"], 'm': d["Peso [Kg]"]*d["Arm [pol]"]} for d in deducoes_lista],
            "additions": [{'desc': a["Descrição"], 'w': a["Peso [Kg]"], 'a': a["Arm [pol]"], 'm': a["Peso [Kg]"]*a["Arm [pol]"] + (momento_extra_flaps if "Flaps" in a["Descrição"] else 0)} for a in adicoes_lista]
        }
        
        excel_data = gerar_excel_por_template(dados_excel, "exemplo_ficha.xlsx")
        st.session_state['excel_data_temp'] = excel_data
        
        st.download_button(
            label="Exportar para Excel (Padrão Oficial)", 
            data=excel_data, 
            file_name=f"{prefixo}_Weighing_Report.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Função auxiliar para mapear as colunas corretas do Banco de Dados
def get_real_col(possible_names):
    if df_historico.empty:
        return possible_names[0]
    for n in possible_names:
        if n in df_historico.columns:
            return n
    return possible_names[0]

def formulario_pesagem(prefixo_selecionado, p_sugerida, r_sugerida, p_anterior, r_anterior, linha_existente=None):
    if linha_existente is None: linha_existente = {}
    info_aero = dict_tipos_aeronave.get(prefixo_selecionado, {})
    tipo_a = info_aero.get('modelo', "")

    aba1, aba2, aba3, aba4, aba5 = st.tabs(["Dados da Ficha", "Células de Carga", "Deductions", "Additions", "Preview"])

    with aba1:
        st.markdown("### Controle de Identificação")
        cp1, cp2 = st.columns(2)
        id_pesagem_input = cp1.number_input("Número da Pesagem:", value=int(p_sugerida), disabled=True)
        rev_input = cp2.number_input("Número da Revisão:", value=int(r_sugerida), disabled=True)
        st.divider()

        c1, c2, c3 = st.columns(3)
        data_emissao = c1.date_input("Data de emissão:", value=datetime.date.today())
        pesado_por = c2.text_input("Pesado por:", value=safe_str(linha_existente.get(get_real_col(['Pesado Por', 'WEIGHED BY']), '')))
        local = c3.text_input("Local da pesagem:", value=safe_str(linha_existente.get(get_real_col(['Local da pesagem']), '')))
        
        c4, c5, c6 = st.columns(3)
        key_data_pes = get_real_col(['Data_da_Pesagem', 'Data da ficha'])
        val_pesagem = pd.to_datetime(linha_existente.get(key_data_pes)).date() if pd.notna(linha_existente.get(key_data_pes)) else datetime.date.today()
        data_pesagem = c4.date_input("Data da pesagem:", value=val_pesagem)
        
        lopa = c5.text_input("LOPA:", value=safe_str(linha_existente.get(get_real_col([' LOPA', 'LOPA']), '')))
        config_lopa = c6.text_input("Configuração LOPA:", value=safe_str(linha_existente.get(get_real_col(['Configuração LOPA ', 'Configuração LOPA']), '')))
        
        c7, c8, c9 = st.columns(3)
        vrbl = c7.text_input("VRBL. NUMBER:", value=safe_str(linha_existente.get(get_real_col(['VRBL', 'VRBL NUMBER']), info_aero.get('vrbl', ''))))
        serial = c8.text_input("SERIAL NUMBER:", value=safe_str(linha_existente.get(get_real_col(['SERIAL', 'SERIAL NUMBER']), info_aero.get('serial', ''))))
        line = c9.text_input("LINE NUMBER:", value=safe_str(linha_existente.get(get_real_col(['LINE', 'LINE NUMBER']), info_aero.get('line', ''))))

        razao = st.text_input("Razão para emissão:", value=safe_str(linha_existente.get(get_real_col(['Motivo', 'Razão']), '')))

    with aba2:
        st.markdown("**Pesagem 01**")
        p1_c1, p1_c2, p1_c3, p1_c4 = st.columns(4)
        p1_nlh = p1_c1.number_input("Nariz LH", value=safe_float(linha_existente.get(get_real_col(['Peso nariz LH']), 0.0)), step=10.0)
        p1_nrh = p1_c2.number_input("Nariz RH", value=safe_float(linha_existente.get(get_real_col(['Peso Nariz RH']), 0.0)), step=10.0)
        p1_rhm = p1_c3.number_input("RH MLG 1", value=safe_float(linha_existente.get(get_real_col(['Peso MLG RH 1 ', 'Peso MLG RH 1']), 0.0)), step=10.0)
        p1_lhm = p1_c4.number_input("LH MLG 1", value=safe_float(linha_existente.get(get_real_col(['Peso MLG LH 1', 'Peso MLG LH 1 ']), 0.0)), step=10.0)
        
        p1_c5, p1_c6, p1_c7, p1_c8 = st.columns(4)
        p1_rhm2 = p1_c7.number_input("RH MLG 2", value=safe_float(linha_existente.get(get_real_col(['Peso MLG RH 2', 'Peso MLG RH 2 ']), 0.0)), step=10.0)
        p1_lhm2 = p1_c8.number_input("LH MLG 2", value=safe_float(linha_existente.get(get_real_col(['Peso MLG LH2', 'Peso MLG LH 2']), 0.0)), step=10.0)

        st.markdown("**Pesagem 02**")
        p2_c1, p2_c2, p2_c3, p2_c4 = st.columns(4)
        p2_nlh = p2_c1.number_input("Nariz LH P2", value=safe_float(linha_existente.get(get_real_col(['Peso nariz LH pesagem 2']), 0.0)), step=10.0)
        p2_nrh = p2_c2.number_input("Nariz RH P2", value=safe_float(linha_existente.get(get_real_col(['Peso Nariz RH pesagem 2']), 0.0)), step=10.0)
        p2_rhm = p2_c3.number_input("RH MLG 1 P2", value=safe_float(linha_existente.get(get_real_col(['Peso MLG RH 1  pesagem 2', 'Peso MLG RH 1 pesagem 2']), 0.0)), step=10.0)
        p2_lhm = p2_c4.number_input("LH MLG 1 P2", value=safe_float(linha_existente.get(get_real_col(['Peso MLG LH 1 pesagem 2']), 0.0)), step=10.0)

        p2_c5, p2_c6, p2_c7, p2_c8 = st.columns(4)
        p2_rhm2 = p2_c7.number_input("RH MLG 2 P2", value=safe_float(linha_existente.get(get_real_col(['Peso MLG RH 2 pesagem 2']), 0.0)), step=10.0)
        p2_lhm2 = p2_c8.number_input("LH MLG 2 P2", value=safe_float(linha_existente.get(get_real_col(['Peso MLG LH2 pesagem 2']), 0.0)), step=10.0)
        
        st.markdown("**Canelas (Inches) e Tail**")
        can_c1, can_c2, can_c3 = st.columns(3)
        canela_lh = can_c1.number_input("Canela LH", value=safe_float(linha_existente.get(get_real_col(['Canela MLG LH']), 0.0)), step=0.1)
        canela_rh = can_c2.number_input("Canela RH", value=safe_float(linha_existente.get(get_real_col(['Canela MLG RH']), 0.0)), step=0.1)
        tail_val = can_c3.number_input("Tail (Kg)", value=safe_float(linha_existente.get(get_real_col(['Tail']), 0.0)), step=10.0)

    with aba3:
        ded_ex = []
        for i in range(1, 16):
            desc = linha_existente.get(get_real_col([f'Deductions description {i}']))
            if pd.notna(desc) and str(desc).strip():
                peso = safe_float(linha_existente.get(get_real_col([f'Deductions Weigth {i}'])))
                arm = safe_float(linha_existente.get(get_real_col([f'Deductions arm {i}'])))
                ded_ex.append({"Descrição": desc, "Peso (Kg)": peso, "Braço (in)": arm, "Momento (kg.in)": peso * arm})
                
        if not ded_ex:
            ded_ex = [{"Descrição": "Fuel (Usable)", "Peso (Kg)": 0.0, "Braço (in)": 660.5, "Momento (kg.in)": 0.0}]
            
        df_ded = pd.DataFrame(ded_ex)
        st.info("💡 **Dica:** O valor visual do Momento na tabela é calculado com os dados iniciais.")
        deducoes_editadas = st.data_editor(
            df_ded, 
            num_rows="dynamic", 
            use_container_width=True,
            column_config={"Momento (kg.in)": st.column_config.NumberColumn(disabled=True)}
        )

    with aba4:
        momento_extra_flaps = 0
        if tipo_a == "B737-MAX": momento_extra_flaps = 6190
        elif tipo_a in ["B737-800", "B737-800SFP", "737-800", "B737-700"]: momento_extra_flaps = 5930

        add_ex = []
        for i in range(1, 17):
            desc = linha_existente.get(get_real_col([f'Additions Description {i}']))
            if pd.notna(desc) and str(desc).strip():
                peso = safe_float(linha_existente.get(get_real_col([f'Additions weigth {i}'])))
                arm = safe_float(linha_existente.get(get_real_col([f'Additions arm {i}'])))
                m_calc = peso * arm
                if "Flaps" in str(desc): m_calc += momento_extra_flaps
                add_ex.append({"Descrição": desc, "Peso (Kg)": peso, "Braço (in)": arm, "Momento (kg.in)": m_calc})
                
        if not add_ex:
            add_ex = [{"Descrição": "Flaps 0 - 40° (up when weighed)", "Peso (Kg)": 0.0, "Braço (in)": 0.0, "Momento (kg.in)": momento_extra_flaps}]

        df_add = pd.DataFrame(add_ex)
        st.info(f"💡 **Dica:** O momento do Flap será exibido automaticamente na coluna 'Momento' sempre que houver a palavra 'Flaps' na descrição.")
        adicoes_editadas = st.data_editor(
            df_add, 
            num_rows="dynamic", 
            use_container_width=True,
            column_config={"Momento (kg.in)": st.column_config.NumberColumn(disabled=True)}
        )

    with aba5:
        lh_val = (p1_lhm + p1_lhm2 + p2_lhm + p2_lhm2) / 2
        rh_val = (p1_rhm + p1_rhm2 + p2_rhm + p2_rhm2) / 2
        nose_val = ((p1_nlh + p2_nlh) / 2) + ((p1_nrh + p2_nrh) / 2)
        
        arm_a = info_aero.get('armnose', 93.0000)
        arm_b_lh = (canela_lh * 0.1304) + 706.17
        arm_b_rh = (canela_rh * 0.1304) + 706.17
        arm_c = 1298.9000

        m_lh = lh_val * arm_b_lh
        m_rh = rh_val * arm_b_rh
        m_nose = nose_val * arm_a
        m_tail = tail_val * arm_c

        tot_reg_w = lh_val + rh_val + nose_val + tail_val
        tot_reg_m = m_lh + m_rh + m_nose + m_tail
        tot_reg_arm = tot_reg_m / tot_reg_w if tot_reg_w > 0 else 0.0

        ded_w = deducoes_editadas["Peso (Kg)"].sum() if not deducoes_editadas.empty else 0.0
        ded_m = (deducoes_editadas["Peso (Kg)"] * deducoes_editadas["Braço (in)"]).sum() if not deducoes_editadas.empty else 0.0
        ded_arm = ded_m / ded_w if ded_w > 0 else 0.0

        add_w = adicoes_editadas["Peso (Kg)"].sum() if not adicoes_editadas.empty else 0.0
        add_m = 0.0
        for idx, row in adicoes_editadas.iterrows():
            m = row["Peso (Kg)"] * row["Braço (in)"]
            if "Flaps" in str(row["Descrição"]): m += momento_extra_flaps
            add_m += m
        add_arm = add_m / add_w if add_w > 0 else 0.0

        basic_w = tot_reg_w + add_w - ded_w
        basic_m = tot_reg_m + add_m - ded_m
        basic_arm = basic_m / basic_w if basic_w > 0 else 0.0
        cg_mac_val = ((basic_arm - 627.1) / 1.558) if basic_arm > 0 else 0.0

        st.dataframe(pd.DataFrame({
            "Reaction / Item": ["LH", "RH", "NOSE", "TAIL", "TOTAL REGISTERED", "DEDUCTIONS", "ADDITIONS", "AIRCRAFT BASIC WEIGHT"],
            "Weight (kg)": [lh_val, rh_val, nose_val, tail_val, tot_reg_w, ded_w, add_w, basic_w],
            "Arm (inch)": [arm_b_lh, arm_b_rh, arm_a, arm_c, tot_reg_arm, ded_arm, add_arm, basic_arm],
            "Moment (kg x inch)": [m_lh, m_rh, m_nose, m_tail, tot_reg_m, ded_m, add_m, basic_m]
        }), use_container_width=True, hide_index=True)
        
        c_res1, c_res2 = st.columns(2)
        c_res1.metric("Aircraft Basic Weight", f"{basic_w:,.4f} kg")
        c_res1.metric("Aircraft Basic Arm", f"{basic_arm:,.4f} in")
        c_res2.metric("C.G. (% MAC)", f"{cg_mac_val:.2f} %")
        
        dados_excel = {
            "prefixo": prefixo_selecionado, "modelo": tipo_a, "pesado_por": pesado_por,
            "local": local, "data": str(data_pesagem), "config_lopa": config_lopa, "lopa": lopa,
            "razao": razao, "cg_mac": cg_mac_val,
            "arm_a": arm_a, "arm_b_lh": arm_b_lh, "arm_b_rh": arm_b_rh,
            "issue_date": str(data_emissao),
            "ultima_pesagem": str(p_anterior) if p_anterior > 0 else "",
            "ultima_revisao": str(r_anterior) if p_anterior > 0 else "",
            "vrbl_number": vrbl, "serial_number": serial, "line_number": line,
            "numero_pesagem": str(int(id_pesagem_input)), "revisao": str(int(rev_input)),
            "reactions": {
                'LH': {'weight': lh_val, 'arm': arm_b_lh, 'moment': m_lh},
                'RH': {'weight': rh_val, 'arm': arm_b_rh, 'moment': m_rh},
                'NOSE': {'weight': nose_val, 'arm': arm_a, 'moment': m_nose},
                'TAIL': {'weight': tail_val, 'arm': arm_c, 'moment': m_tail},
                'TOTAL REGISTERED': {'weight': tot_reg_w, 'arm': tot_reg_arm, 'moment': tot_reg_m},
                'DEDUCTIONS': {'weight': ded_w, 'arm': ded_arm, 'moment': ded_m},
                'ADDITIONS': {'weight': add_w, 'arm': add_arm, 'moment': add_m},
                'AIRCRAFT BASIC WEIGHT': {'weight': basic_w, 'arm': basic_arm, 'moment': basic_m}
            },
            "deductions": [{'desc': d["Descrição"], 'w': d["Peso (Kg)"], 'a': d["Braço (in)"], 'm': d["Peso (Kg)"]*d["Braço (in)"]} for d in deducoes_editadas.to_dict('records')],
            "additions": [{'desc': a["Descrição"], 'w': a["Peso (Kg)"], 'a': a["Braço (in)"], 'm': (a["Peso (Kg)"]*a["Braço (in)"]) + (momento_extra_flaps if "Flaps" in str(a["Descrição"]) else 0)} for a in adicoes_editadas.to_dict('records')]
        }
        
        st.session_state['excel_data_temp'] = gerar_excel_por_template(dados_excel, "exemplo_ficha.xlsx")

    # Mapeamento 100% seguro para evitar o Database Error
    novo_registro = linha_existente.copy()
    
    novo_registro[get_real_col(['Prefixo'])] = prefixo_selecionado
    novo_registro[get_real_col(['Data da ficha', 'Data_da_Pesagem'])] = str(data_emissao)
    novo_registro[get_real_col(['Pesado Por', 'WEIGHED BY'])] = pesado_por
    novo_registro[get_real_col(['Local da pesagem'])] = local
    novo_registro[get_real_col(['Data_da_Pesagem'])] = str(data_pesagem)
    novo_registro[get_real_col([' LOPA', 'LOPA'])] = lopa
    novo_registro[get_real_col(['Configuração LOPA ', 'Configuração LOPA'])] = config_lopa
    novo_registro[get_real_col(['Motivo', 'Razão'])] = razao
    novo_registro[get_real_col(['VRBL', 'VRBL NUMBER'])] = vrbl
    novo_registro[get_real_col(['SERIAL', 'SERIAL NUMBER'])] = serial
    novo_registro[get_real_col(['LINE', 'LINE NUMBER'])] = line
    
    novo_registro[get_real_col(['Peso nariz LH'])] = p1_nlh
    novo_registro[get_real_col(['Peso Nariz RH'])] = p1_nrh
    novo_registro[get_real_col(['Peso MLG RH 1 ', 'Peso MLG RH 1'])] = p1_rhm
    novo_registro[get_real_col(['Peso MLG LH 1', 'Peso MLG LH 1 '])] = p1_lhm
    novo_registro[get_real_col(['Peso MLG RH 2', 'Peso MLG RH 2 '])] = p1_rhm2
    novo_registro[get_real_col(['Peso MLG LH2', 'Peso MLG LH 2'])] = p1_lhm2
    
    novo_registro[get_real_col(['Peso nariz LH pesagem 2'])] = p2_nlh
    novo_registro[get_real_col(['Peso Nariz RH pesagem 2'])] = p2_nrh
    novo_registro[get_real_col(['Peso MLG RH 1  pesagem 2', 'Peso MLG RH 1 pesagem 2'])] = p2_rhm
    novo_registro[get_real_col(['Peso MLG LH 1 pesagem 2'])] = p2_lhm
    novo_registro[get_real_col(['Peso MLG RH 2 pesagem 2'])] = p2_rhm2
    novo_registro[get_real_col(['Peso MLG LH2 pesagem 2'])] = p2_lhm2
    
    novo_registro[get_real_col(['Canela MLG LH'])] = canela_lh
    novo_registro[get_real_col(['Canela MLG RH'])] = canela_rh
    novo_registro[get_real_col(['Tail'])] = tail_val
    
    for i in range(1, 16):
        novo_registro[get_real_col([f'Deductions description {i}'])] = None
        novo_registro[get_real_col([f'Deductions Weigth {i}'])] = None
        novo_registro[get_real_col([f'Deductions arm {i}'])] = None
    for i, row in enumerate(deducoes_editadas.to_dict('records')):
        novo_registro[get_real_col([f'Deductions description {i+1}'])] = row.get('Descrição')
        novo_registro[get_real_col([f'Deductions Weigth {i+1}'])] = row.get('Peso (Kg)')
        novo_registro[get_real_col([f'Deductions arm {i+1}'])] = row.get('Braço (in)')

    for i in range(1, 17):
        novo_registro[get_real_col([f'Additions Description {i}'])] = None
        novo_registro[get_real_col([f'Additions weigth {i}'])] = None
        novo_registro[get_real_col([f'Additions arm {i}'])] = None
    for i, row in enumerate(adicoes_editadas.to_dict('records')):
        novo_registro[get_real_col([f'Additions Description {i+1}'])] = row.get('Descrição')
        novo_registro[get_real_col([f'Additions weigth {i+1}'])] = row.get('Peso (Kg)')
        novo_registro[get_real_col([f'Additions arm {i+1}'])] = row.get('Braço (in)')

    # FILTRO FINAL: Apenas permite salvar chaves que REALMENTE existem no banco de dados SQLite
    chaves_validas = set(df_historico.columns)
    if chaves_validas:
        novo_registro = {k: v for k, v in novo_registro.items() if k in chaves_validas}

    return novo_registro, id_pesagem_input, rev_input

def tela_nova_ficha():
    st.title("Gerar Nova Ficha")
    chaves_validas = [str(k) for k in dict_tipos_aeronave.keys() if str(k) not in ('nan', 'None', '')]
    prefixo = st.selectbox("Aeronave", [""] + sorted(chaves_validas))
    
    if prefixo:
        st.divider()
        df_aero = df_historico[df_historico['Prefixo'] == prefixo].copy()
        
        if not df_aero.empty:
            df_aero['Pesagem_num'] = pd.to_numeric(df_aero['Pesagem'], errors='coerce').fillna(0)
            df_aero['Revisao_num'] = pd.to_numeric(df_aero['Revisao'], errors='coerce').fillna(0)
            df_aero = df_aero.sort_values(by=['Pesagem_num', 'Revisao_num'])
            
            ultima_linha = df_aero.iloc[-1]
            linha_base = ultima_linha.to_dict()
            linha_base.pop('Pesagem_num', None)
            linha_base.pop('Revisao_num', None)
            
            u_pes = int(ultima_linha['Pesagem_num'])
            u_rev = int(ultima_linha['Revisao_num'])
            
            st.info(f"📋 **Base Encontrada:** A última ficha dessa aeronave foi a **Pesagem {u_pes} | Revisão {u_rev}**. Os dados foram copiados para agilizar o preenchimento.")
            tipo_acao = st.radio("Escolha a ação desejada:", ["Nova Revisão (Mantém Pesagem atual)", "Nova Pesagem (Inicia do zero)"], horizontal=True)
            
            if "Revisão" in tipo_acao:
                p_nova = u_pes
                r_nova = u_rev + 1
            else:
                p_nova = u_pes + 1
                r_nova = 0
                
            p_ant = u_pes
            r_ant = u_rev
        else:
            st.info("🆕 **Nenhuma ficha anterior encontrada.** Iniciando a primeira Pesagem (1), Revisão (0) com a ficha em branco.")
            linha_base = {}
            p_nova = 1
            r_nova = 0
            p_ant = 0
            r_ant = 0
            
        novo_registro, p_final, r_final = formulario_pesagem(prefixo, p_nova, r_nova, p_ant, r_ant, linha_existente=linha_base)
        
        st.divider()
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("Salvar Ficha", type="primary", use_container_width=True):
                novo_registro['Pesagem'] = str(int(p_final))
                novo_registro['Revisao'] = str(int(r_final))
                conn = sqlite3.connect('aeronaves.db')
                pd.DataFrame([novo_registro]).to_sql('pesagens', conn, if_exists='append', index=False)
                conn.close()
                st.cache_data.clear()
                st.success("Ficha cadastrada com sucesso.")
        with col_btn2:
            st.download_button(label="Exportar Prévia para Excel", data=st.session_state.get('excel_data_temp', b''), file_name=f"{prefixo}_Preview.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

def tela_edicao():
    st.title("Editar Ficha Existente")
    if st.session_state['nivel_acesso'] != 1:
        st.error("Acesso restrito à Engenharia.")
        return

    prefixos = [""] + sorted(df_historico['Prefixo'].dropna().unique().tolist())
    prefixo = st.selectbox("Aeronave", prefixos)
    
    if prefixo:
        df_filtrado = df_historico[df_historico['Prefixo'] == prefixo]
        pesagem = st.selectbox("Pesagem a editar", [""] + sorted(df_filtrado['Pesagem'].dropna().unique().tolist()))
        
        if pesagem:
            revisao = st.selectbox("Revisão a editar", [""] + sorted(df_filtrado[df_filtrado['Pesagem'] == pesagem]['Revisao'].dropna().unique().tolist()))
            
            if revisao:
                st.divider()
                linha_atual = df_filtrado[(df_filtrado['Pesagem'] == pesagem) & (df_filtrado['Revisao'] == revisao)].iloc[0].to_dict()
                linha_atual.pop('Pesagem_num', None)
                linha_atual.pop('Revisao_num', None)
                
                u_pes = int(safe_float(linha_atual.get('Pesagem', 1)))
                u_rev = int(safe_float(linha_atual.get('Revisao', 0)))
                
                p_nova = u_pes
                r_nova = u_rev + 1
                p_ant = u_pes
                r_ant = u_rev
                
                novo_registro, p_final, r_final = formulario_pesagem(prefixo, p_nova, r_nova, p_ant, r_ant, linha_existente=linha_atual)
                
                st.divider()
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    if st.button("Salvar Nova Revisão", type="primary", use_container_width=True):
                        novo_registro['Pesagem'] = str(int(p_final))
                        novo_registro['Revisao'] = str(int(r_final))
                        conn = sqlite3.connect('aeronaves.db')
                        pd.DataFrame([novo_registro]).to_sql('pesagens', conn, if_exists='append', index=False)
                        conn.close()
                        st.cache_data.clear()
                        st.success("Revisão salva com sucesso.")
                with col_btn2:
                    st.download_button(label="Exportar Revisão para Excel", data=st.session_state.get('excel_data_temp', b''), file_name=f"{prefixo}_Revisao_{r_final}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# 5. ROTEAMENTO E BARRA LATERAL
if not st.session_state['usuario_logado']:
    st.title("Sistema de Pesagem e Balanceamento")
    with st.form("login_form"):
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
        if st.form_submit_button("Acessar"):
            if senha == "123":
                st.session_state['usuario_logado'] = True
                st.session_state['nivel_acesso'] = 1 if usuario == "engenharia" else 2
                st.session_state['nome_usuario'] = "Engenharia GOL" if usuario == "engenharia" else "Consulta"
                st.rerun()
            else:
                st.error("Credenciais inválidas.")
else:
    with st.sidebar:
        st.subheader("Menu Principal")
        st.write(f"Usuário ativo: **{st.session_state['nome_usuario']}**")
        st.divider()
        if st.button("Consultar Fichas", use_container_width=True): st.session_state['pagina_atual'] = 'consulta'
        if st.button("Nova Ficha", use_container_width=True): st.session_state['pagina_atual'] = 'nova_ficha'
        if st.session_state['nivel_acesso'] == 1:
            if st.button("Editar Ficha", use_container_width=True): st.session_state['pagina_atual'] = 'edicao'
        st.divider()
        if st.button("Sair do Sistema", use_container_width=True):
            st.session_state['usuario_logado'] = False
            st.rerun()

    if st.session_state['pagina_atual'] == 'consulta': tela_consulta()
    elif st.session_state['pagina_atual'] == 'nova_ficha': tela_nova_ficha()
    elif st.session_state['pagina_atual'] == 'edicao': tela_edicao()