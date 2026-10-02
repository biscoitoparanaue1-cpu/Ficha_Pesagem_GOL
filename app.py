import io
import os
import sqlite3
import datetime
import hashlib
import hmac
import secrets

import openpyxl
import pandas as pd
import streamlit as st

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


def safe_identifier(val):
    if pd.isna(val):
        return ""
    if pd.api.types.is_number(val) and float(val).is_integer():
        return str(int(val))
    return str(val).strip()


def momento_flaps_do_registro(registro):
    descricoes_flaps = [
        safe_str(registro.get(f"Additions Description {indice}", ""))
        for indice in range(1, 17)
    ]
    if not any("flap" in descricao.casefold() for descricao in descricoes_flaps):
        return None

    momento = safe_float(registro.get("Adction LAP", 0))
    if momento > 0:
        return momento

    for indice, descricao in enumerate(descricoes_flaps, start=1):
        if "flap" in descricao.casefold():
            momento = safe_float(registro.get(f"Additions arm {indice}", 0))
            if momento > 0:
                return momento
    return None


@st.cache_data
def buscar_momento_flaps_banco(prefixo, tipo_aeronave):
    df = carregar_dados_banco()
    if df.empty or "Tipo_aeronave" not in df.columns:
        return 0.0

    df = df.copy()
    coluna_revisao = "Revisao" if "Revisao" in df.columns else "revisao"
    df["_pesagem_num"] = pd.to_numeric(df.get("Pesagem"), errors="coerce").fillna(-1)
    df["_revisao_num"] = pd.to_numeric(df.get(coluna_revisao), errors="coerce").fillna(-1)

    if prefixo and "Prefixo" in df.columns:
        historico_prefixo = df.loc[
            df["Prefixo"].astype(str).str.strip().str.casefold()
            == safe_str(prefixo).casefold()
        ].sort_values(["_pesagem_num", "_revisao_num"], ascending=False)
        for _, registro in historico_prefixo.iterrows():
            momento = momento_flaps_do_registro(registro)
            if momento is not None:
                return momento

    modelo = safe_str(tipo_aeronave).casefold()
    historico_modelo = df.loc[
        df["Tipo_aeronave"].astype(str).str.strip().str.casefold() == modelo
    ]
    momentos = [
        momento
        for _, registro in historico_modelo.iterrows()
        if (momento := momento_flaps_do_registro(registro)) is not None
    ]
    if not momentos:
        return 0.0

    frequencias = pd.Series(momentos).value_counts()
    maior_frequencia = frequencias.max()
    return float(max(frequencias[frequencias == maior_frequencia].index))


def converter_unidade_canela(unidade_key, unidade_anterior_key, chaves_valores):
    unidade_anterior = st.session_state.get(unidade_anterior_key, "in")
    unidade_atual = st.session_state.get(unidade_key, unidade_anterior)
    if unidade_atual != unidade_anterior:
        fator = 25.4 if unidade_anterior == "in" else 1 / 25.4
        for chave in chaves_valores:
            if chave in st.session_state:
                st.session_state[chave] *= fator
        st.session_state[unidade_anterior_key] = unidade_atual


def valor_canela_em_polegadas(valor, unidade):
    return valor / 25.4 if unidade == "mm" else valor


def separar_campos_lopa(registro):
    valores = [
        safe_str(registro.get(" LOPA", registro.get("LOPA", ""))),
        safe_str(registro.get("Configuração LOPA ", registro.get("Configuração LOPA", ""))),
    ]
    valores = [
        valor for valor in valores
        if valor and valor.casefold() not in {"lopa", "configuração lopa"}
    ]
    codigo_lopa = next(
        (valor for valor in valores if valor.upper().startswith("GLP-")), ""
    )
    configuracao = next(
        (valor for valor in valores if not valor.upper().startswith("GLP-")), ""
    )
    return codigo_lopa, configuracao


LEVEL_CORRECTION_VALUES = {
    "-2": 3.0,
    "-1 7/8": 2.8,
    "-1 3/4": 2.7,
    "-1 5/8": 2.5,
    "-1 1/2": 2.3,
    "-1 3/8": 2.1,
    "-1 1/4": 1.9,
    "-1 1/8": 1.7,
    "-1": 1.5,
    "-7/8": 1.3,
    "-3/4": 1.1,
    "-5/8": 0.9,
    "-1/2": 0.8,
    "-3/8": 0.6,
    "-1/4": 0.4,
    "-1/8": 0.2,
    "0": 0.0,
    "1/8": -0.2,
    "1/4": -0.4,
    "3/8": -0.6,
    "1/2": -0.8,
    "5/8": -0.9,
    "3/4": -1.1,
    "7/8": -1.3,
    "1": -1.5,
    "1 1/8": -1.7,
    "1 1/4": -1.9,
    "1 3/8": -2.1,
    "1 1/2": -2.3,
    "1 5/8": -2.5,
    "1 3/4": -2.7,
    "1 7/8": -2.8,
    "2": -3.0,
}
LEVEL_CORRECTION_LABEL = "Level Correction (Value added to CG)"


def normalizar_angulo_level_correction(angulo):
    angulo = safe_str(angulo).replace("\xa0", " ")
    angulo = " ".join(angulo.split())
    if angulo.endswith(".0") and angulo[:-2] in LEVEL_CORRECTION_VALUES:
        angulo = angulo[:-2]
    return angulo if angulo in LEVEL_CORRECTION_VALUES else ""


def calcular_level_correction(angulo, peso_base):
    angulo = normalizar_angulo_level_correction(angulo)
    if not angulo:
        return 0.0, 0.0, None

    valor_correspondente = LEVEL_CORRECTION_VALUES[angulo]
    if angulo.startswith("-"):
        return valor_correspondente, valor_correspondente * peso_base, "additions"
    return valor_correspondente, -valor_correspondente * peso_base, "deductions"

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
        nomes_colunas = {
            coluna: " ".join(str(coluna).replace("\n", " ").split()).upper()
            for coluna in df_tipos.columns
        }
        col_serial = next(
            (
                coluna for coluna, nome in nomes_colunas.items()
                if nome in {"S/N", "SERIAL", "SERIAL NUMBER"}
            ),
            None,
        )
        if col_serial is None:
            col_serial = next(
                (
                    coluna for coluna, nome in nomes_colunas.items()
                    if "SERIAL" in nome and "ENGINE" not in nome
                ),
                None,
            )
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
                    'vrbl': safe_identifier(row[col_vrbl]) if col_vrbl else "",
                    'serial': safe_identifier(row[col_serial]) if col_serial else "",
                    'line': safe_identifier(row[col_line]) if col_line else ""
                }
        return dict_aero
    except Exception as e:
        return {}

df_historico = carregar_dados_banco()
dict_tipos_aeronave = carregar_tipos_aeronave()

def obter_senha_admin_inicial():
    senha = os.environ.get("INITIAL_ADMIN_PASSWORD")
    if senha:
        return senha
    try:
        return st.secrets.get("INITIAL_ADMIN_PASSWORD")
    except Exception:
        return None


def inicializar_controle_acesso():
    with sqlite3.connect('aeronaves.db') as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS usuarios (
                usuario TEXT PRIMARY KEY COLLATE NOCASE,
                nome TEXT NOT NULL,
                senha_hash TEXT NOT NULL,
                salt TEXT NOT NULL,
                nivel_acesso INTEGER NOT NULL,
                criado_em TEXT NOT NULL
            )
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS fluxo_fichas (
                prefixo TEXT NOT NULL,
                pesagem TEXT NOT NULL,
                revisao TEXT NOT NULL,
                gerado_por TEXT,
                aprovado_por TEXT,
                criado_em TEXT NOT NULL,
                aprovado_em TEXT,
                PRIMARY KEY (prefixo, pesagem, revisao)
            )
        ''')
        usuario_existente = conn.execute('SELECT 1 FROM usuarios LIMIT 1').fetchone()
        if not usuario_existente:
            senha_inicial = obter_senha_admin_inicial()
            if not senha_inicial:
                st.error(
                    "Configure INITIAL_ADMIN_PASSWORD em Streamlit Secrets "
                    "para criar o primeiro usuário de Engenharia."
                )
            else:
                salt = secrets.token_hex(16)
                senha_hash = hashlib.pbkdf2_hmac(
                    'sha256', senha_inicial.encode('utf-8'), bytes.fromhex(salt), 600000
                ).hex()
                conn.execute(
                    '''INSERT INTO usuarios
                       (usuario, nome, senha_hash, salt, nivel_acesso, criado_em)
                       VALUES (?, ?, ?, ?, ?, ?)''',
                    ('engenharia', 'Engenharia GOL', senha_hash, salt, 1,
                     datetime.datetime.now(datetime.timezone.utc).isoformat())
                )
def autenticar_usuario(usuario, senha):
    with sqlite3.connect('aeronaves.db') as conn:
        conn.row_factory = sqlite3.Row
        registro = conn.execute(
            'SELECT * FROM usuarios WHERE usuario = ?', (usuario.strip(),)
        ).fetchone()
    if not registro:
        return None
    senha_hash = hashlib.pbkdf2_hmac(
        'sha256', senha.encode('utf-8'), bytes.fromhex(registro['salt']), 600000
    ).hex()
    if not hmac.compare_digest(senha_hash, registro['senha_hash']):
        senha_inicial = obter_senha_admin_inicial()
        if (
            registro['usuario'].casefold() != 'engenharia'
            or not senha_inicial
            or not hmac.compare_digest(senha, senha_inicial)
        ):
            return None

        salt = secrets.token_hex(16)
        senha_hash = hashlib.pbkdf2_hmac(
            'sha256', senha_inicial.encode('utf-8'), bytes.fromhex(salt), 600000
        ).hex()
        with sqlite3.connect('aeronaves.db') as conn:
            conn.execute(
                'UPDATE usuarios SET senha_hash = ?, salt = ? WHERE usuario = ?',
                (senha_hash, salt, registro['usuario']),
            )
        registro = dict(registro)
        registro['senha_hash'] = senha_hash
        registro['salt'] = salt
    return dict(registro)

def criar_usuario(nome, usuario, senha, nivel_acesso):
    salt = secrets.token_hex(16)
    senha_hash = hashlib.pbkdf2_hmac(
        'sha256', senha.encode('utf-8'), bytes.fromhex(salt), 600000
    ).hex()
    with sqlite3.connect('aeronaves.db') as conn:
        conn.execute(
            '''INSERT INTO usuarios
               (usuario, nome, senha_hash, salt, nivel_acesso, criado_em)
               VALUES (?, ?, ?, ?, ?, ?)''',
            (usuario.strip(), nome.strip(), senha_hash, salt, nivel_acesso,
             datetime.datetime.now(datetime.timezone.utc).isoformat())
        )

def registrar_ficha(prefixo, pesagem, revisao, usuario):
    agora = datetime.datetime.now(datetime.timezone.utc).isoformat()
    with sqlite3.connect('aeronaves.db') as conn:
        conn.execute(
            '''INSERT INTO fluxo_fichas
               (prefixo, pesagem, revisao, gerado_por, criado_em)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT(prefixo, pesagem, revisao) DO UPDATE SET
                   gerado_por = excluded.gerado_por,
                   aprovado_por = NULL,
                   criado_em = excluded.criado_em,
                   aprovado_em = NULL''',
            (safe_str(prefixo), safe_str(pesagem), safe_str(revisao), usuario, agora)
        )

def buscar_fluxo_ficha(prefixo, pesagem, revisao):
    with sqlite3.connect('aeronaves.db') as conn:
        conn.row_factory = sqlite3.Row
        registro = conn.execute(
            '''SELECT fluxo_fichas.gerado_por AS gerador_usuario,
                      gerador.nome AS gerador_nome,
                      fluxo_fichas.aprovado_por AS aprovador_usuario,
                      aprovador.nome AS aprovador_nome,
                      fluxo_fichas.aprovado_em
               FROM fluxo_fichas
               LEFT JOIN usuarios AS gerador
                   ON gerador.usuario = fluxo_fichas.gerado_por
               LEFT JOIN usuarios AS aprovador
                   ON aprovador.usuario = fluxo_fichas.aprovado_por
               WHERE prefixo = ? AND pesagem = ? AND revisao = ?''',
            (safe_str(prefixo), safe_str(pesagem), safe_str(revisao))
        ).fetchone()
    if not registro:
        return {
            'gerador_usuario': None,
            'gerador_nome': 'Não registrado',
            'aprovador_nome': 'Pendente',
            'aprovado_em': None,
        }
    return {
        'gerador_usuario': registro['gerador_usuario'],
        'gerador_nome': registro['gerador_nome'] or 'Não registrado',
        'aprovador_nome': registro['aprovador_nome'] or 'Pendente',
        'aprovado_em': registro['aprovado_em'],
    }

def carregar_fluxos_fichas():
    with sqlite3.connect('aeronaves.db') as conn:
        conn.row_factory = sqlite3.Row
        registros = conn.execute(
            '''SELECT fluxo_fichas.prefixo, fluxo_fichas.pesagem,
                      fluxo_fichas.revisao,
                      fluxo_fichas.gerado_por AS gerador_usuario,
                      gerador.nome AS gerador_nome,
                      fluxo_fichas.aprovado_por AS aprovador_usuario,
                      aprovador.nome AS aprovador_nome
               FROM fluxo_fichas
               LEFT JOIN usuarios AS gerador
                   ON gerador.usuario = fluxo_fichas.gerado_por
               LEFT JOIN usuarios AS aprovador
                   ON aprovador.usuario = fluxo_fichas.aprovado_por'''
        ).fetchall()
    return {
        (safe_str(registro['prefixo']), safe_str(registro['pesagem']), safe_str(registro['revisao'])): {
            'gerador_usuario': registro['gerador_usuario'],
            'gerador_nome': registro['gerador_nome'] or 'Não registrado',
            'aprovador_usuario': registro['aprovador_usuario'],
            'aprovador_nome': registro['aprovador_nome'] or 'Pendente',
        }
        for registro in registros
    }

def aprovar_ficha(prefixo, pesagem, revisao, usuario):
    agora = datetime.datetime.now(datetime.timezone.utc).isoformat()
    with sqlite3.connect('aeronaves.db') as conn:
        conn.execute(
            '''INSERT INTO fluxo_fichas
               (prefixo, pesagem, revisao, aprovado_por, criado_em, aprovado_em)
               VALUES (?, ?, ?, ?, ?, ?)
               ON CONFLICT(prefixo, pesagem, revisao) DO UPDATE SET
                   aprovado_por = excluded.aprovado_por,
                   aprovado_em = excluded.aprovado_em''',
            (safe_str(prefixo), safe_str(pesagem), safe_str(revisao), usuario, agora, agora)
        )

inicializar_controle_acesso()

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
    fluxo = buscar_fluxo_ficha(prefixo, pesagem, revisao)
    st.caption(f"Gerada por: {fluxo['gerador_nome']} | Aprovada por: {fluxo['aprovador_nome']}")
    aba1, aba2, aba3, aba4, aba5 = st.tabs(["Dados Gerais", "Células de Carga", "Deductions", "Additions", "Weighing Report"])
    info_aero = dict_tipos_aeronave.get(prefixo, {})
    lopa, config_lopa = separar_campos_lopa(row)

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
        c5.text_input("LOPA:", value=lopa, disabled=True)
        c6.text_input("Configuração LOPA:", value=config_lopa, disabled=True)
        
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
            rh_1_p2 = safe_float(row.get('Peso MLG RH 1  pesagem 2', row.get('Peso MLG RH 1 pesagem 2', 0)))
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

        level_correction_angle = normalizar_angulo_level_correction(
            row.get('Graus correção do cg', '')
        )
        level_correction_factor, level_correction_moment, level_correction_side = (
            calcular_level_correction(level_correction_angle, tot_reg_weight)
        )

        ded_weight = sum([item["Peso [Kg]"] for item in deducoes_lista])
        ded_moment = sum([item["Peso [Kg]"] * item["Arm [pol]"] for item in deducoes_lista])
        if level_correction_side == "deductions":
            ded_moment += level_correction_moment
        ded_arm = ded_moment / ded_weight if ded_weight > 0 else 0.0

        flaps_ativo = any("flap" in str(item.get("Descrição", "")).casefold() for item in adicoes_lista)
        momento_extra_flaps = 0.0
        if flaps_ativo:
            momento_extra_flaps = (
                momento_flaps_do_registro(row)
                or buscar_momento_flaps_banco(prefixo, tipo_aeronave)
            )

        add_weight = sum([item["Peso [Kg]"] for item in adicoes_lista])
        add_moment = sum([item["Peso [Kg]"] * item["Arm [pol]"] for item in adicoes_lista]) + momento_extra_flaps
        if level_correction_side == "additions":
            add_moment += level_correction_moment
        add_arm = add_moment / add_weight if add_weight > 0 else 0.0

        basic_weight = tot_reg_weight + add_weight - ded_weight
        basic_moment = tot_reg_moment + add_moment - ded_moment
        basic_arm = basic_moment / basic_weight if basic_weight > 0 else 0.0
        cg_mac = ((basic_arm - 627.1) / 1.558) if basic_arm > 0 else 0.0

        report_data = {
            "Reaction / Item": ["LH", "RH", "NOSE", "TAIL", "TOTAL REGISTERED", "LEVEL CORRECTION", "DEDUCTIONS", "ADDITIONS", "AIRCRAFT BASIC WEIGHT"],
            "Weight (kg)": [lh, rh, nose, tail, tot_reg_weight, 0.0, ded_weight, add_weight, basic_weight],
            "Arm (inch)": [arm_b_lh, arm_b_rh, arm_a, arm_c, tot_reg_arm, level_correction_factor, ded_arm, add_arm, basic_arm],
            "Moment (kg x inch)": [m_lh, m_rh, m_nose, m_tail, tot_reg_moment, level_correction_moment, ded_moment, add_moment, basic_moment]
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
        lopa, config_lopa = separar_campos_lopa(row)

        deducoes_excel = [
            {'desc': d["Descrição"], 'w': d["Peso [Kg]"], 'a': d["Arm [pol]"], 'm': d["Peso [Kg]"] * d["Arm [pol]"]}
            for d in deducoes_lista
        ]
        adicoes_excel = [
            {'desc': a["Descrição"], 'w': a["Peso [Kg]"], 'a': a["Arm [pol]"], 'm': a["Peso [Kg]"] * a["Arm [pol]"] + (momento_extra_flaps if "flap" in a["Descrição"].casefold() else 0)}
            for a in adicoes_lista
        ]
        if level_correction_side == "deductions":
            deducoes_excel.append({
                'desc': LEVEL_CORRECTION_LABEL, 'w': 0.0,
                'a': level_correction_factor, 'm': level_correction_moment,
            })
        elif level_correction_side == "additions":
            adicoes_excel.append({
                'desc': LEVEL_CORRECTION_LABEL, 'w': 0.0,
                'a': level_correction_factor, 'm': level_correction_moment,
            })

        dados_excel = {
            "prefixo": prefixo, "modelo": tipo_aeronave, "pesado_por": safe_str(row.get('Pesado Por', '')),
            "local": safe_str(row.get('Local da pesagem', '')), "data": safe_str(row.get('Data_da_Pesagem', '')),
            "config_lopa": config_lopa, "lopa": lopa,
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
            "deductions": deducoes_excel,
            "additions": adicoes_excel
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

PRESETS = {
    "motivo": [
        "5 Years Check",
        "IFE Installation",
        "New Painting",
        "New Production Aircraft",
        "Scheduled",
        "Seats Reconfiguration",
    ],
    "pesado_por": ["GOL", "Boeing", "DIGEX", "Flightstar", "Aeroman", "Transavia"],
    "local": [
        "Lagoa Santa, MG - Brazil",
        "Renton, WA - EUA",
        "Seattle, WA - EUA",
        "Jacksonville, FL - EUA",
        "São José dos Campos, SP - Brazil",
        "San Salvador - El Salvador",
        "Amsterdam - Netherlands",
        "Everett, WA - EUA",
    ],
    "config_lopa": [],
    "lopa": [],
    "deductions": [
        "Down Lock, Nose Gear",
        "Down Lock, Main Gear",
        "Plumb Bob",
        "In-seat Power Removal",
        "Seats Removal (177 Pax)",
        "Seats Removal (138 Pax)",
        "Tail Jack Adaptors",
        "Wing Jack Adaptors",
        "Carbon Brake Retrofit Program (SB 737-32-1429)",
        "Radial Tires Replacement (SB 737-32-1535)",
        "Flaps 0 - 40° (Up when weighed)",
        "Wiring Diagrams",
        "Production Flight Emergency Equipment",
        "Protective Carpet Runner",
        "Plastic Seat Cover",
        "Loto Kit",
        "Leather Seats Cover Installation",
        "Seats Removal (168 Pax)",
    ],
    "additions": [
        "2Ku System Antenna Installation (SB 737-44-1016)",
        "2Ku System Cabin Equip Installation (SB 737-44-1017)",
        "Carbon Brake Retrofit Program (SB 737-32-1429)",
        "Container, Oxygen Mask Spares",
        "Disable Person Kit Aft (1 EA)",
        "Dual navAero EFB System Installation",
        "Emergency Care Kit (1 EA)",
        "Emergency Medical Kit (1 EA)",
        "First Aid Kit Aft (2 EA)",
        "First Aid Kit Fwd (1 EA)",
        "Flaps 0 - 40° (Up when weighed)",
        "Fuel, Drain, Unusable (5.6 Gal)",
        "Fuel, Trapped, Unusable (14.2 Gal)",
        "Fuel, Trapped, Usable (5.5 Gal)",
        "Fwd Cargo, Aft Bulkhead Panels",
        "In-Seat Power Supply System (ISPSS) Installation",
        "Seats Installation (138 Pax)",
        "Seats Installation (186 Pax)",
        "Seats Installation (177 Pax)",
        "Lavatory Dry Supplies",
        "Lavatory Pre-charge Fluid (6.0 Gal)",
        "Life Vest Installation",
        "Waste Container",
        "Nitrogen Generation System Install. (SB 737-47-1003)",
        "Oil, Auxiliary Power Unit (2.3 Gal)",
        "Oil, Drain, Unusable Engine (6.6 Gal)",
        "Oil, Drain, Usable Engine (10.2 Gal)",
        "Oil, Integrated Drive Gen (4.4 Gal)",
        "Oil, Trapped, Unusable Engine (1.0 Gal)",
        "Water, Potable (62.6 Gal)",
        "Water, Potable (42.2 Gal)",
        "Water, Potable (52.2 Gal)",
        "Oven Tray (8 EA)",
        "Oxigen Mask, Disposable (2 EA)",
        "Oxigen Mask, Disposable (3 EA)",
        "Oxygen Access Panel",
        "Oxygen Bottle, Portable (2 EA)",
        "Oxygen Cilinder Assy",
        "Oxygen Cilinder Charge (18 EA)",
        "Sat LINK Communication Management System",
        "Jungle Survival Kit Aft (1 EA)",
        "Jungle Survival Kit Aft (2 EA)",
        "Jungle Survival Kit Fwd (2 EA)",
        "Securelink/ MQAR System",
        "Standard Container",
        "Toillet Chemicals",
        "Operations Manual",
        "Oven",
        "Oven Rack",
        "Wheelchair (1 EA)",
    ],
}

valores_lopa_historico = set()
for coluna in (" LOPA", "LOPA", "Configuração LOPA ", "Configuração LOPA"):
    if coluna in df_historico.columns:
        valores_lopa_historico.update(
            safe_str(valor)
            for valor in df_historico[coluna].dropna().tolist()
            if safe_str(valor).casefold() not in {"lopa", "configuração lopa"}
        )

PRESETS["lopa"] = sorted(
    {
        valor for valor in valores_lopa_historico
        if valor.upper().startswith("GLP-")
    } | {"GLP-MAX8-001-XMC"},
    key=str.casefold,
)
PRESETS["config_lopa"] = sorted(
    {
        valor for valor in valores_lopa_historico
        if not valor.upper().startswith("GLP-")
    } | {"186 Pax + 10 Flight Crew"},
    key=str.casefold,
)


def campo_com_preset(container, label, valor, opcoes, key):
    valor = safe_str(valor)
    escolhas = [""] + list(opcoes)
    if valor and valor not in escolhas:
        escolhas.insert(1, valor)
    return container.selectbox(
        label,
        escolhas,
        index=escolhas.index(valor) if valor else 0,
        key=key,
        accept_new_options=True,
        placeholder="Digite ou selecione",
    )


def renderizar_editor_itens(container, itens, opcoes, key, momento_flaps=0):
    colunas = ["Descrição", "Peso (Kg)", "Braço (in)", "Momento (kg.in)"]
    quantidade_inicial = max(len(itens), 1)
    chave_linhas = f"{key}_linhas"
    chave_proximo_id = f"{key}_proximo_id"
    if chave_linhas not in st.session_state:
        quantidade_anterior = max(
            quantidade_inicial,
            st.session_state.get(f"{key}_quantidade", quantidade_inicial),
        )
        st.session_state[chave_linhas] = [str(indice) for indice in range(quantidade_anterior)]
        st.session_state[chave_proximo_id] = quantidade_anterior

    ids_linhas = list(st.session_state[chave_linhas])
    cabecalho = container.columns([4, 1, 1, 1, 0.9])
    for coluna, titulo in zip(cabecalho, [*colunas, "Ação"]):
        coluna.caption(titulo)

    linhas = []
    for posicao, id_linha in enumerate(ids_linhas):
        indice_item = int(id_linha) if id_linha.isdigit() else len(itens)
        item = itens[indice_item] if indice_item < len(itens) else {}
        descricao_inicial = safe_str(item.get("Descrição", ""))
        escolhas = [""] + list(opcoes)
        if descricao_inicial and descricao_inicial not in escolhas:
            escolhas.append(descricao_inicial)

        descricao_col, peso_col, braco_col, momento_col, excluir_col = container.columns([4, 1, 1, 1, 0.9])
        descricao = descricao_col.selectbox(
            f"Descrição, linha {posicao + 1}",
            escolhas,
            index=escolhas.index(descricao_inicial),
            key=f"{key}_{id_linha}_descricao",
            accept_new_options=True,
            placeholder="Digite ou selecione",
            label_visibility="collapsed",
        )
        peso = peso_col.number_input(
            f"Peso (Kg), linha {posicao + 1}",
            value=safe_float(item.get("Peso (Kg)", 0.0)),
            step=0.1,
            key=f"{key}_{id_linha}_peso",
            label_visibility="collapsed",
        )
        braco = braco_col.number_input(
            f"Braço (in), linha {posicao + 1}",
            value=safe_float(item.get("Braço (in)", 0.0)),
            step=0.1,
            key=f"{key}_{id_linha}_braco",
            label_visibility="collapsed",
        )

        descricao = safe_str(descricao)
        momento = peso * braco
        if momento_flaps and "flap" in descricao.casefold():
            momento += momento_flaps
        momento_col.write(f"{momento:,.2f}")

        if descricao:
            linhas.append({
                "Descrição": descricao,
                "Peso (Kg)": peso,
                "Braço (in)": braco,
                "Momento (kg.in)": momento,
            })

        excluir_col.button(
            "",
            key=f"{key}_{id_linha}_excluir",
            help="Excluir esta linha",
            on_click=excluir_linha,
            args=(chave_linhas, key, id_linha),
            type="tertiary",
            icon=":material/delete:",
        )

    adicionar_col, _ = container.columns(2)
    adicionar_col.button(
        "Adicionar item",
        key=f"{key}_adicionar",
        on_click=adicionar_linha,
        args=(chave_linhas, chave_proximo_id),
    )
    return pd.DataFrame(linhas, columns=colunas)


def adicionar_linha(chave_linhas, chave_proximo_id):
    novo_id = st.session_state[chave_proximo_id]
    st.session_state[chave_linhas] = [
        *st.session_state[chave_linhas], f"extra_{novo_id}"
    ]
    st.session_state[chave_proximo_id] = novo_id + 1


def excluir_linha(chave_linhas, prefixo_widget, id_linha):
    st.session_state[chave_linhas] = [
        linha for linha in st.session_state[chave_linhas] if linha != id_linha
    ]
    for campo in ("descricao", "peso", "braco"):
        st.session_state.pop(f"{prefixo_widget}_{id_linha}_{campo}", None)

def formulario_pesagem(prefixo_selecionado, p_sugerida, r_sugerida, p_anterior, r_anterior, linha_existente=None):
    if linha_existente is None: linha_existente = {}
    info_aero = dict_tipos_aeronave.get(prefixo_selecionado, {})
    tipo_a = info_aero.get('modelo', "")
    form_key = f"{prefixo_selecionado}_{p_sugerida}_{r_sugerida}"

    aba1, aba2, aba3, aba4, aba5, aba6 = st.tabs(["Dados da Ficha", "Células de Carga", "Deductions", "Additions", "Level Correction", "Preview"])

    with aba1:
        st.markdown("### Controle de Identificação")
        cp1, cp2 = st.columns(2)
        id_pesagem_input = cp1.number_input("Número da Pesagem:", value=int(p_sugerida), disabled=True)
        rev_input = cp2.number_input("Número da Revisão:", value=int(r_sugerida), disabled=True)
        st.divider()

        c1, c2, c3 = st.columns(3)
        data_emissao = c1.date_input("Data de emissão:", value=datetime.date.today())
        pesado_por = campo_com_preset(c2, "Pesado por:", linha_existente.get(get_real_col(['Pesado Por', 'WEIGHED BY']), ''), PRESETS["pesado_por"], f"{form_key}_pesado_por")
        local = campo_com_preset(c3, "Local da pesagem:", linha_existente.get(get_real_col(['Local da pesagem']), ''), PRESETS["local"], f"{form_key}_local")
        
        c4, c5, c6 = st.columns(3)
        key_data_pes = get_real_col(['Data_da_Pesagem', 'Data da ficha'])
        val_pesagem = pd.to_datetime(linha_existente.get(key_data_pes)).date() if pd.notna(linha_existente.get(key_data_pes)) else datetime.date.today()
        data_pesagem = c4.date_input("Data da pesagem:", value=val_pesagem)
        
        lopa_inicial, config_lopa_inicial = separar_campos_lopa(linha_existente)
        lopa = campo_com_preset(c5, "LOPA:", lopa_inicial, PRESETS["lopa"], f"{form_key}_lopa")
        config_lopa = campo_com_preset(c6, "Configuração LOPA:", config_lopa_inicial, PRESETS["config_lopa"], f"{form_key}_config_lopa")
        
        c7, c8, c9 = st.columns(3)
        vrbl = c7.text_input("VRBL. NUMBER:", value=safe_str(linha_existente.get(get_real_col(['VRBL', 'VRBL NUMBER']), info_aero.get('vrbl', ''))))
        serial = c8.text_input("SERIAL NUMBER:", value=safe_str(linha_existente.get(get_real_col(['SERIAL', 'SERIAL NUMBER']), info_aero.get('serial', ''))))
        line = c9.text_input("LINE NUMBER:", value=safe_str(linha_existente.get(get_real_col(['LINE', 'LINE NUMBER']), info_aero.get('line', ''))))

        razao = campo_com_preset(st, "Razão para emissão:", linha_existente.get(get_real_col(['Motivo', 'Razão']), ''), PRESETS["motivo"], f"{form_key}_motivo")

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
        
        st.markdown("**Canelas**")
        unidade_canela_key = f"{form_key}_unidade_canela"
        unidade_canela_anterior_key = f"{form_key}_unidade_canela_anterior"
        chave_canela_lh = f"{form_key}_canela_lh"
        chave_canela_rh = f"{form_key}_canela_rh"
        st.session_state.setdefault(unidade_canela_key, "in")
        st.session_state.setdefault(
            unidade_canela_anterior_key, st.session_state[unidade_canela_key]
        )
        fator_unidade = 25.4 if st.session_state[unidade_canela_key] == "mm" else 1.0
        st.session_state.setdefault(
            chave_canela_lh,
            safe_float(linha_existente.get(get_real_col(['Canela MLG LH']), 0.0)) * fator_unidade,
        )
        st.session_state.setdefault(
            chave_canela_rh,
            safe_float(linha_existente.get(get_real_col(['Canela MLG RH']), 0.0)) * fator_unidade,
        )
        unidade_canela = st.radio(
            "Unidade da canela",
            ["in", "mm"],
            horizontal=True,
            key=unidade_canela_key,
            on_change=converter_unidade_canela,
            args=(
                unidade_canela_key,
                unidade_canela_anterior_key,
                (chave_canela_lh, chave_canela_rh),
            ),
        )
        can_c1, can_c2 = st.columns(2)
        canela_lh_input = can_c1.number_input(
            f"Canela LH ({unidade_canela})",
            step=0.1 if unidade_canela == "in" else 1.0,
            key=chave_canela_lh,
        )
        canela_rh_input = can_c2.number_input(
            f"Canela RH ({unidade_canela})",
            step=0.1 if unidade_canela == "in" else 1.0,
            key=chave_canela_rh,
        )
        canela_lh = valor_canela_em_polegadas(canela_lh_input, unidade_canela)
        canela_rh = valor_canela_em_polegadas(canela_rh_input, unidade_canela)
        tail_val = 0.0

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

        deducoes_editadas = renderizar_editor_itens(
            st, ded_ex, PRESETS["deductions"], f"{form_key}_deductions"
        )

    with aba4:
        momento_extra_flaps = momento_flaps_do_registro(linha_existente)
        if momento_extra_flaps is None:
            momento_extra_flaps = buscar_momento_flaps_banco(
                prefixo_selecionado, tipo_a
            )

        add_ex = []
        for i in range(1, 17):
            desc = linha_existente.get(get_real_col([f'Additions Description {i}']))
            if pd.notna(desc) and str(desc).strip():
                peso = safe_float(linha_existente.get(get_real_col([f'Additions weigth {i}'])))
                arm = safe_float(linha_existente.get(get_real_col([f'Additions arm {i}'])))
                m_calc = peso * arm
                if "flap" in str(desc).casefold(): m_calc += momento_extra_flaps
                add_ex.append({"Descrição": desc, "Peso (Kg)": peso, "Braço (in)": arm, "Momento (kg.in)": m_calc})
                
        if not add_ex:
            add_ex = [{"Descrição": "Flaps 0 - 40° (up when weighed)", "Peso (Kg)": 0.0, "Braço (in)": 0.0, "Momento (kg.in)": momento_extra_flaps}]

        adicoes_editadas = renderizar_editor_itens(
            st,
            add_ex,
            PRESETS["additions"],
            f"{form_key}_additions",
            momento_flaps=momento_extra_flaps,
        )

    with aba5:
        angulo_salvo = normalizar_angulo_level_correction(
            linha_existente.get(get_real_col(['Graus correção do cg']), '')
        )
        opcoes_angulo = [""] + list(LEVEL_CORRECTION_VALUES)
        angulo_level_correction = st.selectbox(
            "Ângulo de inclinação",
            opcoes_angulo,
            index=opcoes_angulo.index(angulo_salvo),
            key=f"{form_key}_level_correction",
        )
        fator_level_correction = LEVEL_CORRECTION_VALUES.get(angulo_level_correction, 0.0)
        st.metric("Valor correspondente", f"{fator_level_correction:.1f}")

    with aba6:
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
        level_correction_factor, level_correction_moment, level_correction_side = (
            calcular_level_correction(angulo_level_correction, tot_reg_w)
        )
        if level_correction_side == "deductions":
            ded_m += level_correction_moment
        ded_arm = ded_m / ded_w if ded_w > 0 else 0.0

        add_w = adicoes_editadas["Peso (Kg)"].sum() if not adicoes_editadas.empty else 0.0
        add_m = 0.0
        for idx, row in adicoes_editadas.iterrows():
            m = row["Peso (Kg)"] * row["Braço (in)"]
            if "flap" in str(row["Descrição"]).casefold(): m += momento_extra_flaps
            add_m += m
        if level_correction_side == "additions":
            add_m += level_correction_moment
        add_arm = add_m / add_w if add_w > 0 else 0.0

        basic_w = tot_reg_w + add_w - ded_w
        basic_m = tot_reg_m + add_m - ded_m
        basic_arm = basic_m / basic_w if basic_w > 0 else 0.0
        cg_mac_val = ((basic_arm - 627.1) / 1.558) if basic_arm > 0 else 0.0

        st.dataframe(pd.DataFrame({
            "Reaction / Item": ["LH", "RH", "NOSE", "TAIL", "TOTAL REGISTERED", "LEVEL CORRECTION", "DEDUCTIONS", "ADDITIONS", "AIRCRAFT BASIC WEIGHT"],
            "Weight (kg)": [lh_val, rh_val, nose_val, tail_val, tot_reg_w, 0.0, ded_w, add_w, basic_w],
            "Arm (inch)": [arm_b_lh, arm_b_rh, arm_a, arm_c, tot_reg_arm, level_correction_factor, ded_arm, add_arm, basic_arm],
            "Moment (kg x inch)": [m_lh, m_rh, m_nose, m_tail, tot_reg_m, level_correction_moment, ded_m, add_m, basic_m]
        }), use_container_width=True, hide_index=True)
        
        c_res1, c_res2 = st.columns(2)
        c_res1.metric("Aircraft Basic Weight", f"{basic_w:,.4f} kg")
        c_res1.metric("Aircraft Basic Arm", f"{basic_arm:,.4f} in")
        c_res2.metric("C.G. (% MAC)", f"{cg_mac_val:.2f} %")
        
        deducoes_excel = [
            {'desc': d["Descrição"], 'w': d["Peso (Kg)"], 'a': d["Braço (in)"], 'm': d["Peso (Kg)"] * d["Braço (in)"]}
            for d in deducoes_editadas.to_dict('records')
        ]
        adicoes_excel = [
            {'desc': a["Descrição"], 'w': a["Peso (Kg)"], 'a': a["Braço (in)"], 'm': a["Peso (Kg)"] * a["Braço (in)"] + (momento_extra_flaps if "flap" in str(a["Descrição"]).casefold() else 0)}
            for a in adicoes_editadas.to_dict('records')
        ]
        if level_correction_side == "deductions":
            deducoes_excel.append({
                'desc': LEVEL_CORRECTION_LABEL, 'w': 0.0,
                'a': level_correction_factor, 'm': level_correction_moment,
            })
        elif level_correction_side == "additions":
            adicoes_excel.append({
                'desc': LEVEL_CORRECTION_LABEL, 'w': 0.0,
                'a': level_correction_factor, 'm': level_correction_moment,
            })

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
            "deductions": deducoes_excel,
            "additions": adicoes_excel
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
    novo_registro[get_real_col(['Adction LAP'])] = (
        momento_extra_flaps
        if any("flap" in str(descricao).casefold() for descricao in adicoes_editadas["Descrição"])
        else 0.0
    )
    novo_registro[get_real_col(['Graus correção do cg'])] = angulo_level_correction or None
    novo_registro[get_real_col(['Braço correspondente '])] = (
        LEVEL_CORRECTION_VALUES[angulo_level_correction]
        if angulo_level_correction else None
    )
    
    for i in range(1, 16):
        novo_registro[get_real_col([f'Deductions description {i}'])] = None
        novo_registro[get_real_col([f'Deductions Weigth {i}'])] = None
        novo_registro[get_real_col([f'Deductions arm {i}'])] = None
    slots_deducoes = [
        i for i in range(1, 16)
        if not len(df_historico.columns) or all(
            coluna in df_historico.columns
            for coluna in (
                f'Deductions description {i}',
                f'Deductions Weigth {i}',
                f'Deductions arm {i}',
            )
        )
    ]
    for i, row in zip(slots_deducoes, deducoes_editadas.to_dict('records')):
        novo_registro[get_real_col([f'Deductions description {i}'])] = row.get('Descrição')
        novo_registro[get_real_col([f'Deductions Weigth {i}'])] = row.get('Peso (Kg)')
        novo_registro[get_real_col([f'Deductions arm {i}'])] = row.get('Braço (in)')

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

        if st.checkbox(
            "Definir número da Pesagem e Revisão manualmente",
            key=f"numeracao_manual_{prefixo}",
        ):
            col_pesagem, col_revisao = st.columns(2)
            p_nova = col_pesagem.number_input(
                "Número da Pesagem",
                min_value=1,
                value=max(1, int(p_nova)),
                step=1,
                key=f"pesagem_manual_{prefixo}",
            )
            r_nova = col_revisao.number_input(
                "Número da Revisão",
                min_value=0,
                value=max(0, int(r_nova)),
                step=1,
                key=f"revisao_manual_{prefixo}",
            )
            
        novo_registro, p_final, r_final = formulario_pesagem(prefixo, p_nova, r_nova, p_ant, r_ant, linha_existente=linha_base)
        
        st.divider()
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("Salvar Ficha", type="primary", use_container_width=True):
                ficha_existente = False
                if not df_aero.empty:
                    ficha_existente = (
                        (df_aero['Pesagem_num'] == int(p_final))
                        & (df_aero['Revisao_num'] == int(r_final))
                    ).any()

                if ficha_existente:
                    st.error(f"Já existe uma ficha para {prefixo}, Pesagem {int(p_final)}, Revisão {int(r_final)}.")
                else:
                    novo_registro['Pesagem'] = str(int(p_final))
                    novo_registro['Revisao'] = str(int(r_final))
                    conn = sqlite3.connect('aeronaves.db')
                    pd.DataFrame([novo_registro]).to_sql('pesagens', conn, if_exists='append', index=False)
                    conn.close()
                    registrar_ficha(prefixo, p_final, r_final, st.session_state['usuario_id'])
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
                        registrar_ficha(prefixo, p_final, r_final, st.session_state['usuario_id'])
                        st.cache_data.clear()
                        st.success("Revisão salva com sucesso.")
                with col_btn2:
                    st.download_button(label="Exportar Revisão para Excel", data=st.session_state.get('excel_data_temp', b''), file_name=f"{prefixo}_Revisao_{r_final}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

def tela_criar_login():
    if st.session_state['nivel_acesso'] != 1:
        st.error("Acesso restrito à Engenharia.")
        return

    st.title("Criar login")
    with st.form("form_criar_login"):
        nome = st.text_input("Nome completo")
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
        confirmar_senha = st.text_input("Confirmar senha", type="password")
        nivel = st.selectbox("Perfil", ["Consulta", "Engenharia"])
        enviar = st.form_submit_button("Criar usuário", type="primary")

    if enviar:
        if not nome.strip() or not usuario.strip():
            st.error("Informe o nome e o usuário.")
        elif len(senha) < 1:
            st.error("A senha deve ter pelo menos 1 caracteres.")
        elif senha != confirmar_senha:
            st.error("As senhas não coincidem.")
        else:
            try:
                criar_usuario(nome, usuario, senha, 1 if nivel == "Engenharia" else 2)
                st.success(f"Login criado para {nome.strip()}.")
            except sqlite3.IntegrityError:
                st.error("Esse nome de usuário já está cadastrado.")

def tela_aprovar_fichas():
    if st.session_state['nivel_acesso'] != 1:
        st.error("Acesso restrito à Engenharia.")
        return

    st.title("Aprovar ficha de pesagem")
    colunas_chave = ['Prefixo', 'Pesagem', 'Revisao']
    if df_historico.empty or not all(coluna in df_historico.columns for coluna in colunas_chave):
        st.info("Não há fichas disponíveis para aprovação.")
        return

    prefixos = sorted(df_historico['Prefixo'].dropna().astype(str).unique().tolist())
    prefixo_selecionado = st.selectbox("Aeronave", prefixos)
    fichas_aeronave = df_historico.loc[
        df_historico['Prefixo'].astype(str) == prefixo_selecionado,
        colunas_chave,
    ].drop_duplicates()
    fluxos = carregar_fluxos_fichas()
    opcoes = {}
    for _, ficha in fichas_aeronave.iterrows():
        prefixo = safe_str(ficha['Prefixo'])
        pesagem = safe_str(ficha['Pesagem'])
        revisao = safe_str(ficha['Revisao'])
        fluxo = fluxos.get((prefixo, pesagem, revisao), {'aprovador_usuario': None})
        if not fluxo['aprovador_usuario']:
            rotulo = f"{prefixo} | Pesagem {pesagem} | Revisão {revisao}"
            opcoes[rotulo] = (prefixo, pesagem, revisao)

    if not opcoes:
        st.success("Todas as fichas desta aeronave estão aprovadas.")
        return

    selecionada = st.selectbox("Ficha pendente", list(opcoes))
    prefixo, pesagem, revisao = opcoes[selecionada]
    fluxo = buscar_fluxo_ficha(prefixo, pesagem, revisao)
    st.write(f"Gerada por: {fluxo['gerador_nome']}")
    if fluxo['gerador_usuario'] == st.session_state['usuario_id']:
        st.warning("Quem gerou a ficha não pode aprová-la.")
    elif st.button("Aprovar ficha", type="primary"):
        aprovar_ficha(prefixo, pesagem, revisao, st.session_state['usuario_id'])
        st.success("Ficha aprovada.")
        st.rerun()

def excluir_ficha(prefixo, pesagem, revisao):
    with sqlite3.connect('aeronaves.db') as conn:
        colunas = {registro[1] for registro in conn.execute('PRAGMA table_info(pesagens)')}
        coluna_revisao = 'Revisao' if 'Revisao' in colunas else 'revisao'
        if not {'Prefixo', 'Pesagem', coluna_revisao}.issubset(colunas):
            raise sqlite3.OperationalError('As colunas de identificação da ficha não foram encontradas.')

        cursor = conn.execute(
            f'''DELETE FROM pesagens
                WHERE CAST(Prefixo AS TEXT) = ?
                  AND CAST(Pesagem AS TEXT) = ?
                  AND CAST("{coluna_revisao}" AS TEXT) = ?''',
            (safe_str(prefixo), safe_str(pesagem), safe_str(revisao)),
        )
        quantidade = cursor.rowcount
        if quantidade:
            conn.execute(
                '''DELETE FROM fluxo_fichas
                   WHERE prefixo = ? AND pesagem = ? AND revisao = ?''',
                (safe_str(prefixo), safe_str(pesagem), safe_str(revisao)),
            )
    return quantidade


def tela_excluir_ficha():
    if st.session_state['nivel_acesso'] != 1:
        st.error("Acesso restrito à Engenharia.")
        return

    st.title("Excluir ficha")
    colunas_chave = ['Prefixo', 'Pesagem', 'Revisao']
    if df_historico.empty or not all(coluna in df_historico.columns for coluna in colunas_chave):
        st.info("Não há fichas disponíveis para exclusão.")
        return

    prefixos = sorted(df_historico['Prefixo'].dropna().astype(str).unique().tolist())
    prefixo = st.selectbox("Aeronave", prefixos, key="excluir_prefixo")
    fichas = df_historico.loc[
        df_historico['Prefixo'].astype(str) == prefixo,
        ['Pesagem', 'Revisao'],
    ].dropna().astype(str).drop_duplicates()
    opcoes = {
        f"Pesagem {linha.Pesagem} | Revisão {linha.Revisao}": (linha.Pesagem, linha.Revisao)
        for linha in fichas.itertuples(index=False)
    }
    if not opcoes:
        st.info("Não há fichas para esta aeronave.")
        return

    selecionada = st.selectbox("Ficha", list(opcoes), key="excluir_ficha")
    pesagem, revisao = opcoes[selecionada]
    fluxo = buscar_fluxo_ficha(prefixo, pesagem, revisao)
    st.write(f"Gerada por: {fluxo['gerador_nome']}")
    st.write(f"Aprovada por: {fluxo['aprovador_nome']}")

    with st.form("form_excluir_ficha"):
        st.warning("A exclusão é permanente e também remove os dados de aprovação desta ficha.")
        confirmar = st.checkbox(
            f"Confirmo excluir {prefixo}, Pesagem {pesagem}, Revisão {revisao}."
        )
        enviar = st.form_submit_button("Excluir ficha", type="primary")

    if enviar:
        if not confirmar:
            st.error("Marque a confirmação para excluir a ficha.")
        else:
            removidas = excluir_ficha(prefixo, pesagem, revisao)
            if removidas:
                st.cache_data.clear()
                st.success(f"Ficha excluída. Registros removidos: {removidas}.")
                st.rerun()
            else:
                st.error("A ficha não foi encontrada no banco de dados.")

# 5. ROTEAMENTO E BARRA LATERAL
if not st.session_state['usuario_logado']:
    st.title("Sistema de Pesagem e Balanceamento")
    with st.form("login_form"):
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
        if st.form_submit_button("Acessar"):
            registro_usuario = autenticar_usuario(usuario, senha)
            if registro_usuario:
                st.session_state['usuario_logado'] = True
                st.session_state['usuario_id'] = registro_usuario['usuario']
                st.session_state['nivel_acesso'] = registro_usuario['nivel_acesso']
                st.session_state['nome_usuario'] = registro_usuario['nome']
                st.session_state['pagina_atual'] = 'consulta'
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")
else:
    with st.sidebar:
        st.subheader("Menu Principal")
        st.write(f"Usuário ativo: **{st.session_state['nome_usuario']}**")
        st.divider()
        if st.button("Consultar Fichas", use_container_width=True): st.session_state['pagina_atual'] = 'consulta'
        if st.button("Nova Ficha", use_container_width=True): st.session_state['pagina_atual'] = 'nova_ficha'
        if st.session_state['nivel_acesso'] == 1:
            if st.button("Editar Ficha", use_container_width=True): st.session_state['pagina_atual'] = 'edicao'
            if st.button("Aprovar Ficha", use_container_width=True): st.session_state['pagina_atual'] = 'aprovar'
            if st.button("Criar login", use_container_width=True): st.session_state['pagina_atual'] = 'criar_login'
            if st.button("Excluir ficha", use_container_width=True): st.session_state['pagina_atual'] = 'excluir'
        st.divider()
        if st.button("Sair do Sistema", use_container_width=True):
            st.session_state['usuario_logado'] = False
            st.session_state['nivel_acesso'] = 0
            st.session_state['usuario_id'] = ""
            st.session_state['nome_usuario'] = ""
            st.session_state['pagina_atual'] = 'consulta'
            st.rerun()

    if st.session_state['pagina_atual'] == 'consulta': tela_consulta()
    elif st.session_state['pagina_atual'] == 'nova_ficha': tela_nova_ficha()
    elif st.session_state['pagina_atual'] == 'edicao': tela_edicao()
    elif st.session_state['pagina_atual'] == 'aprovar': tela_aprovar_fichas()
    elif st.session_state['pagina_atual'] == 'criar_login': tela_criar_login()
    elif st.session_state['pagina_atual'] == 'excluir': tela_excluir_ficha()