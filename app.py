import io
import os
import datetime
import functools
import hashlib
import hmac
import json
import secrets

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PillowImage, UnidentifiedImageError
from sqlalchemy import text
from sqlalchemy.engine import make_url
from sqlalchemy.exc import DBAPIError, IntegrityError

# 1. CONFIGURAÇÃO INICIAL
st.set_page_config(page_title="Pesagem e Balanceamento", layout="wide", initial_sidebar_state="expanded")

if 'usuario_logado' not in st.session_state:
    st.session_state['usuario_logado'] = False
    st.session_state['nivel_acesso'] = 0 
    st.session_state['nome_usuario'] = ""
if 'pagina_atual' not in st.session_state:
    st.session_state['pagina_atual'] = 'consulta'

st.markdown(
    """
    <style>
    [data-testid="stToolbar"], [data-testid="stDecoration"],
    [data-testid="stAppDeployButton"] { display: none !important; }
    header[data-testid="stHeader"] { background: transparent; }
    .block-container { padding-top: 2.2rem; max-width: 1280px; }
    section[data-testid="stSidebar"] {
        background: #FFFFFF; border-right: 1px solid #E6E8EC;
    }
    section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: .35rem; }
    section[data-testid="stSidebar"] hr { margin: .9rem 0; }
    input:disabled, textarea:disabled {
        color: #1D2433 !important; -webkit-text-fill-color: #1D2433 !important;
        opacity: 1 !important; cursor: default;
    }
    [data-testid="stWidgetLabel"] p { color: #1D2433 !important; opacity: 1 !important; }
    [data-baseweb="input"]:has(input:disabled) { background: #FFFFFF !important; }
    section[data-testid="stSidebar"] .stButton button {
        justify-content: flex-start; border: none; box-shadow: none;
        padding: 0.4rem 0.75rem; min-height: 2.3rem; font-weight: 500;
    }
    section[data-testid="stSidebar"] .stButton button > div {
        justify-content: flex-start; width: 100%;
    }
    section[data-testid="stSidebar"] .stButton button[kind="secondary"] {
        background: transparent; color: #3B4352;
    }
    section[data-testid="stSidebar"] .stButton button[kind="secondary"]:hover {
        background: #FFF1E8; color: #FF6A13;
    }
    .marca { display: flex; align-items: center; gap: .6rem; margin-bottom: .4rem; }
    .marca-logo {
        width: 38px; height: 38px; border-radius: 10px; background: #FF6A13;
        color: #fff; display: flex; align-items: center; justify-content: center;
        font-weight: 700; font-size: .8rem; flex-shrink: 0;
    }
    .marca-nome { font-weight: 700; font-size: 1.02rem; line-height: 1.1; }
    .marca-sub { color: #6B7280; font-size: .78rem; }
    .usuario {
        background: #F6F7F9; border-radius: 10px; padding: .55rem .75rem;
        font-size: .85rem; color: #3B4352; margin: 1rem 0 .3rem;
    }
    .grupo-menu {
        color: #9AA1AD; font-size: .72rem; font-weight: 600; letter-spacing: .06em;
        text-transform: uppercase; margin: 1.3rem 0 .45rem .2rem;
    }
    .cabecalho h1 { font-size: 1.75rem; font-weight: 700; margin: 0; padding: 0; }
    .cabecalho p { color: #6B7280; margin: .25rem 0 1.2rem; }
    div[data-testid="stMetric"] {
        background: #FFFFFF; border: 1px solid #E6E8EC; border-radius: 12px;
        padding: .85rem 1rem;
    }
    div[data-testid="stMetricLabel"] { color: #6B7280; }
    .stTabs [data-baseweb="tab-list"] { gap: .25rem; }
    .stTabs [data-baseweb="tab"] { padding: .4rem .9rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


NIVEL_CONSULTA, NIVEL_OPERACAO, NIVEL_ADMIN = 2, 1, 3
PERFIS = {
    NIVEL_CONSULTA: "Consulta",
    NIVEL_OPERACAO: "Gera e aprova",
    NIVEL_ADMIN: "Administrador",
}


def nivel_atual():
    return st.session_state.get('nivel_acesso', 0)


def exigir_nivel(*niveis):
    if nivel_atual() in niveis:
        return True
    st.error("Seu perfil não tem acesso a esta tela.")
    return False


def cabecalho(titulo, subtitulo=""):
    st.markdown(
        f'<div class="cabecalho"><h1>{titulo}</h1>'
        + (f"<p>{subtitulo}</p>" if subtitulo else "<p></p>")
        + "</div>",
        unsafe_allow_html=True,
    )

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


def normalizar_chave_ficha(valor):
    valor = safe_str(valor)
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return valor
    return str(int(numero)) if numero.is_integer() else format(numero, "g")


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
    df = carregar_fichas_para_momento_flaps(prefixo, tipo_aeronave)
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

def protegido(tela):
    """Mostra uma mensagem clara em vez de derrubar o app se algo falhar."""
    @functools.wraps(tela)
    def executar(*args, **kwargs):
        try:
            return tela(*args, **kwargs)
        except Exception as erro:
            st.cache_data.clear()
            st.error(
                "Algo deu errado nesta tela e nada foi perdido do que já estava "
                "salvo. Tente novamente; se o erro continuar, envie o detalhe "
                "abaixo para a Engenharia."
            )
            with st.expander("Detalhe do erro"):
                st.exception(erro)
    return executar


# 2. CONEXÃO E CARREGAMENTO DOS BANCOS DE DADOS
@st.cache_resource
def obter_conexao_postgresql():
    configuracao = st.secrets.get("connections", {}).get("postgresql", {})
    url = configuracao.get("url") if configuracao else None
    if url:
        try:
            make_url(url)
        except ValueError:
            raise ValueError(
                "A URL em [connections.postgresql].url é inválida. "
                "Confira se está no formato esperado e se a porta após o host "
                "é numérica, como 5432 ou 6543."
            ) from None
    # prepare_threshold=None evita erros de prepared statement no pooler.
    # Conexões derrubadas pelo Supabase são refeitas em ConexaoPostgreSQL.execute,
    # sem o custo de um "ping" antes de cada consulta.
    return st.connection(
        "postgresql",
        type="sql",
        pool_recycle=240,
        pool_size=5,
        max_overflow=5,
        connect_args={"prepare_threshold": None},
    )


class LinhaBanco:
    def __init__(self, linha):
        self._valores = tuple(linha)
        self._mapeamento = dict(linha._mapping)

    def __getitem__(self, chave):
        if isinstance(chave, int):
            return self._valores[chave]
        return self._mapeamento[chave]

    def __iter__(self):
        return iter(self._valores)

    def keys(self):
        return self._mapeamento.keys()


class ResultadoBanco:
    def __init__(self, resultado):
        self._resultado = resultado
        self.rowcount = resultado.rowcount

    def fetchone(self):
        linha = self._resultado.fetchone()
        return LinhaBanco(linha) if linha is not None else None

    def fetchall(self):
        return [LinhaBanco(linha) for linha in self._resultado.fetchall()]

    def __iter__(self):
        return iter(self.fetchall())


class ConexaoPostgreSQL:
    def __init__(self):
        self._sessao = obter_conexao_postgresql().session
        self._primeira_consulta = True

    @property
    def row_factory(self):
        return None

    @row_factory.setter
    def row_factory(self, _valor):
        pass

    @staticmethod
    def _preparar(sql, parametros):
        if parametros is None or isinstance(parametros, dict):
            return sql, parametros or {}
        partes = sql.split("?")
        if len(partes) - 1 != len(parametros):
            raise ValueError("A quantidade de parâmetros SQL não corresponde aos placeholders.")
        consulta = partes[0]
        valores = {}
        for indice, valor in enumerate(parametros):
            nome = f"param_{indice}"
            consulta += f":{nome}{partes[indice + 1]}"
            valores[nome] = valor
        return consulta, valores

    def execute(self, sql, parametros=None):
        consulta, valores = self._preparar(sql, parametros)
        primeira = self._primeira_consulta
        self._primeira_consulta = False
        try:
            return ResultadoBanco(self._sessao.execute(text(consulta), valores))
        except DBAPIError as erro:
            # Conexão encerrada pelo servidor: refaz uma vez se nada foi
            # executado ainda nesta transação.
            if not (primeira and erro.connection_invalidated):
                raise
            self._sessao.rollback()
            self._sessao.close()
            self._sessao = obter_conexao_postgresql().session
            return ResultadoBanco(self._sessao.execute(text(consulta), valores))

    def executemany(self, sql, sequencias):
        if not sequencias:
            return None
        consulta, _ = self._preparar(sql, sequencias[0])
        valores = [self._preparar(sql, linha)[1] for linha in sequencias]
        return ResultadoBanco(self._sessao.execute(text(consulta), valores))

    def __enter__(self):
        return self

    def __exit__(self, tipo_erro, _erro, _rastreio):
        try:
            if tipo_erro is None:
                self._sessao.commit()
            else:
                self._sessao.rollback()
        finally:
            self._sessao.close()

    def close(self):
        self._sessao.commit()
        self._sessao.close()


def conectar_banco():
    return ConexaoPostgreSQL()


def normalizar_valor_json(valor):
    if valor is None:
        return None
    if isinstance(valor, dict):
        return {
            str(chave): normalizar_valor_json(item)
            for chave, item in valor.items()
        }
    if isinstance(valor, (list, tuple)):
        return [normalizar_valor_json(item) for item in valor]
    if hasattr(valor, "item"):
        return normalizar_valor_json(valor.item())
    if pd.isna(valor):
        return None
    if isinstance(valor, (datetime.date, datetime.datetime)):
        return valor.isoformat()
    return valor


def preparar_registro_json(registro):
    dados = dict(registro)
    revisao = dados.pop("revisao", None)
    if safe_str(dados.get("Revisao", "")) == "":
        dados["Revisao"] = revisao
    for chave in ("Pesagem", "Revisao"):
        dados[chave] = normalizar_chave_ficha(dados.get(chave, ""))
    dados["Prefixo"] = safe_str(dados.get("Prefixo", ""))
    dados.pop("Pesagem_num", None)
    dados.pop("Revisao_num", None)
    return json.dumps(
        normalizar_valor_json(dados),
        ensure_ascii=False,
        allow_nan=False,
    )


def buscar_id_ficha(conn, prefixo, pesagem, revisao):
    registros = conn.execute(
        """SELECT id, dados->>'Pesagem' AS pesagem,
                  COALESCE(dados->>'Revisao', dados->>'revisao') AS revisao
           FROM pesagens
           WHERE lower(btrim(dados->>'Prefixo')) = lower(btrim(?))
           ORDER BY id DESC""",
        (safe_str(prefixo),),
    ).fetchall()
    return [
        item["id"]
        for item in registros
        if normalizar_chave_ficha(item["pesagem"]) == normalizar_chave_ficha(pesagem)
        and normalizar_chave_ficha(item["revisao"]) == normalizar_chave_ficha(revisao)
    ]


def bloquear_prefixo(conn, prefixo):
    # Serializa gravações da mesma aeronave para evitar fichas duplicadas
    # quando o botão é clicado duas vezes ou duas pessoas salvam juntas.
    conn.execute(
        "SELECT pg_advisory_xact_lock(hashtext(lower(btrim(?))))",
        (safe_str(prefixo),),
    )


def inserir_ficha(registro):
    with conectar_banco() as conn:
        bloquear_prefixo(conn, registro.get("Prefixo", ""))
        if buscar_id_ficha(
            conn,
            registro.get("Prefixo", ""),
            registro.get("Pesagem", ""),
            registro.get("Revisao", registro.get("revisao", "")),
        ):
            return False
        conn.execute(
            "INSERT INTO pesagens (dados) VALUES (CAST(? AS JSONB))",
            (preparar_registro_json(registro),),
        )
    invalidar_cache_fichas()
    return True


def atualizar_ficha(prefixo, pesagem, revisao, registro):
    with conectar_banco() as conn:
        bloquear_prefixo(conn, prefixo)
        ids = buscar_id_ficha(conn, prefixo, pesagem, revisao)
        if not ids:
            return False

        resultado = conn.execute(
            "UPDATE pesagens SET dados = CAST(? AS JSONB) WHERE id = ?",
            (preparar_registro_json(registro), ids[0]),
        )
    atualizada = resultado.rowcount == 1
    if atualizada:
        invalidar_cache_fichas()
    return atualizada


@st.cache_data
def carregar_dados_banco():
    with conectar_banco() as conn:
        registros = conn.execute(
            """SELECT dados FROM pesagens
               WHERE NULLIF(btrim(dados->>'Prefixo'), '') IS NOT NULL
               ORDER BY id"""
        ).fetchall()
    return dataframe_fichas(registros)


def dataframe_fichas(registros):
    df = (
        pd.DataFrame([registro[0] for registro in registros])
        if registros
        else pd.DataFrame(columns=["Prefixo", "Pesagem", "Revisao"])
    )
    if 'revisao' in df.columns:
        if 'Revisao' in df.columns:
            df['Revisao'] = df['Revisao'].where(
                df['Revisao'].map(safe_str) != "", df['revisao']
            )
            df = df.drop(columns=['revisao'])
        else:
            df = df.rename(columns={'revisao': 'Revisao'})
    for coluna in ('Pesagem', 'Revisao'):
        if coluna in df.columns:
            df[coluna] = df[coluna].map(normalizar_chave_ficha)
    if 'Prefixo' in df.columns:
        df = df.loc[df['Prefixo'].map(safe_str) != ""].copy()
        df['Prefixo'] = df['Prefixo'].map(safe_str)
    return df.reset_index(drop=True)


@st.cache_data
def carregar_prefixos_fichas():
    with conectar_banco() as conn:
        registros = conn.execute(
            """SELECT DISTINCT dados->>'Prefixo' AS prefixo
               FROM pesagens
               WHERE NULLIF(btrim(dados->>'Prefixo'), '') IS NOT NULL
               ORDER BY prefixo"""
        ).fetchall()
    return [safe_str(registro["prefixo"]) for registro in registros]


@st.cache_data
def carregar_fichas_prefixo(prefixo):
    return dataframe_fichas(consultar_registros_prefixo(prefixo))


def consultar_registros_prefixo(prefixo):
    with conectar_banco() as conn:
        return conn.execute(
            """SELECT dados FROM pesagens
               WHERE lower(btrim(dados->>'Prefixo')) = lower(btrim(?))
               ORDER BY id""",
            (safe_str(prefixo),),
        ).fetchall()


@st.cache_data
def carregar_linha_ficha(prefixo, pesagem, revisao):
    df = carregar_fichas_prefixo(prefixo)
    if not {"Pesagem", "Revisao"}.issubset(df.columns):
        return None
    linhas = df.loc[
        (df["Pesagem"].map(normalizar_chave_ficha)
         == normalizar_chave_ficha(pesagem))
        & (df["Revisao"].map(normalizar_chave_ficha)
           == normalizar_chave_ficha(revisao))
    ]
    return linhas.iloc[0] if not linhas.empty else None


@st.cache_data
def carregar_fichas_para_momento_flaps(prefixo, tipo_aeronave):
    with conectar_banco() as conn:
        registros = conn.execute(
            """SELECT dados FROM pesagens
               WHERE lower(btrim(dados->>'Prefixo')) = lower(btrim(?))
                  OR lower(btrim(dados->>'Tipo_aeronave')) = lower(btrim(?))
               ORDER BY id DESC""",
            (safe_str(prefixo), safe_str(tipo_aeronave)),
        ).fetchall()
    return dataframe_fichas(registros)


def invalidar_cache_fichas():
    carregar_dados_banco.clear()
    carregar_linha_ficha.clear()
    listar_fichas_pendentes.clear()
    carregar_presets_lopa.clear()
    carregar_prefixos_fichas.clear()
    carregar_fichas_prefixo.clear()
    carregar_fichas_para_momento_flaps.clear()
    buscar_momento_flaps_banco.clear()
    buscar_fluxo_ficha.clear()
    carregar_campos_com_erro.clear()
    carregar_assinatura_usuario.clear()
    listar_fichas_devolvidas_usuario.clear()


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

dict_tipos_aeronave = carregar_tipos_aeronave()


def normalizar_prefixo(prefixo):
    return "".join(safe_str(prefixo).upper().split())


@st.cache_data
def carregar_aeronaves_cadastradas():
    with conectar_banco() as conn:
        registros = conn.execute(
            """SELECT prefixo, modelo, serial, vrbl, line, armnose,
                      cadastrado_por, cadastrado_em
               FROM aeronaves ORDER BY prefixo"""
        ).fetchall()
    return {
        registro['prefixo']: {
            'modelo': registro['modelo'],
            'armnose': (
                float(registro['armnose'])
                if registro['armnose'] is not None
                else armnose_padrao_modelo(registro['modelo'])
            ),
            'vrbl': safe_str(registro['vrbl']),
            'serial': safe_str(registro['serial']),
            'line': safe_str(registro['line']),
            'cadastrado_por': safe_str(registro['cadastrado_por']),
            'cadastrado_em': safe_str(registro['cadastrado_em']),
        }
        for registro in registros
    }


def armnose_padrao_modelo(modelo):
    valores = [
        dados['armnose']
        for dados in dict_tipos_aeronave.values()
        if dados.get('modelo') == modelo
    ]
    if not valores:
        return 93.0
    return float(pd.Series(valores).mode().max())


def cadastrar_aeronave(prefixo, modelo, serial, vrbl, line, armnose, usuario):
    with conectar_banco() as conn:
        conn.execute(
            """INSERT INTO aeronaves
               (prefixo, modelo, serial, vrbl, line, armnose,
                cadastrado_por, cadastrado_em)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (prefixo, modelo, serial, vrbl, line, float(armnose), usuario,
             datetime.datetime.now(datetime.timezone.utc).isoformat()),
        )
    carregar_aeronaves_cadastradas.clear()


def remover_aeronave_cadastrada(prefixo):
    with conectar_banco() as conn:
        conn.execute('DELETE FROM aeronaves WHERE prefixo = ?', (prefixo,))
    carregar_aeronaves_cadastradas.clear()


def obter_senha_admin_inicial():
    senha = os.environ.get("INITIAL_ADMIN_PASSWORD")
    if senha:
        return senha
    try:
        return st.secrets.get("INITIAL_ADMIN_PASSWORD")
    except Exception:
        return None


@st.cache_resource
def inicializar_controle_acesso():
    with conectar_banco() as conn:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS pesagens (
                id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
                dados JSONB NOT NULL
            )
        ''')
        conn.execute('''
            CREATE INDEX IF NOT EXISTS pesagens_prefixo_idx
            ON pesagens (lower(btrim(dados->>'Prefixo')))
        ''')
        conn.execute('''
            CREATE INDEX IF NOT EXISTS pesagens_tipo_aeronave_idx
            ON pesagens (lower(btrim(dados->>'Tipo_aeronave')))
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS usuarios (
                usuario TEXT PRIMARY KEY,
                nome TEXT NOT NULL,
                senha_hash TEXT NOT NULL,
                salt TEXT NOT NULL,
                nivel_acesso INTEGER NOT NULL,
                criado_em TEXT NOT NULL
            )
        ''')
        conn.execute('''
            CREATE UNIQUE INDEX IF NOT EXISTS usuarios_usuario_casefold_uidx
            ON usuarios (lower(usuario))
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
        conn.execute('''
            CREATE TABLE IF NOT EXISTS campos_com_erro (
                prefixo TEXT NOT NULL,
                pesagem TEXT NOT NULL,
                revisao TEXT NOT NULL,
                campo TEXT NOT NULL,
                descricao TEXT NOT NULL,
                marcado_por TEXT,
                marcado_em TEXT NOT NULL,
                PRIMARY KEY (prefixo, pesagem, revisao, campo)
            )
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS opcoes_cadastradas (
                categoria TEXT NOT NULL,
                valor TEXT NOT NULL,
                cadastrado_por TEXT,
                cadastrado_em TEXT NOT NULL,
                PRIMARY KEY (categoria, valor)
            )
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS aeronaves (
                prefixo TEXT PRIMARY KEY,
                modelo TEXT NOT NULL,
                serial TEXT,
                vrbl TEXT,
                line TEXT,
                armnose DOUBLE PRECISION NOT NULL,
                cadastrado_por TEXT,
                cadastrado_em TEXT NOT NULL
            )
        ''')
        # Bancos que já tinham a tabela da versão anterior (só prefixo, modelo,
        # serial, vrbl e line) recebem as colunas novas.
        conn.execute('''
            ALTER TABLE aeronaves
                ADD COLUMN IF NOT EXISTS serial TEXT,
                ADD COLUMN IF NOT EXISTS vrbl TEXT,
                ADD COLUMN IF NOT EXISTS line TEXT,
                ADD COLUMN IF NOT EXISTS armnose DOUBLE PRECISION,
                ADD COLUMN IF NOT EXISTS cadastrado_por TEXT,
                ADD COLUMN IF NOT EXISTS cadastrado_em TEXT
        ''')
        conn.execute('''
            CREATE TABLE IF NOT EXISTS assinaturas_usuarios (
                usuario TEXT PRIMARY KEY,
                imagem_png BYTEA NOT NULL,
                atualizado_em TEXT NOT NULL,
                FOREIGN KEY (usuario) REFERENCES usuarios(usuario) ON DELETE CASCADE
            )
        ''')
        # A conta inicial da Engenharia passa a ser Administrador (nível 3).
        conn.execute(
            "UPDATE usuarios SET nivel_acesso = 3 WHERE usuario = 'engenharia' AND nivel_acesso = 1"
        )
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
                    ('engenharia', 'Engenharia GOL', senha_hash, salt, 3,
                     datetime.datetime.now(datetime.timezone.utc).isoformat())
                )
def autenticar_usuario(usuario, senha):
    with conectar_banco() as conn:
        registro = conn.execute(
            'SELECT * FROM usuarios WHERE lower(usuario) = lower(?)',
            (usuario.strip(),),
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
        with conectar_banco() as conn:
            conn.execute(
                'UPDATE usuarios SET senha_hash = ?, salt = ? WHERE usuario = ?',
                (senha_hash, salt, registro['usuario']),
            )
        invalidar_cache_fichas()
        registro = dict(registro)
        registro['senha_hash'] = senha_hash
        registro['salt'] = salt
    return dict(registro)

def criar_usuario(nome, usuario, senha, nivel_acesso):
    salt = secrets.token_hex(16)
    senha_hash = hashlib.pbkdf2_hmac(
        'sha256', senha.encode('utf-8'), bytes.fromhex(salt), 600000
    ).hex()
    with conectar_banco() as conn:
        conn.execute(
            '''INSERT INTO usuarios
               (usuario, nome, senha_hash, salt, nivel_acesso, criado_em)
               VALUES (?, ?, ?, ?, ?, ?)''',
            (usuario.strip(), nome.strip(), senha_hash, salt, nivel_acesso,
             datetime.datetime.now(datetime.timezone.utc).isoformat())
        )
    invalidar_cache_fichas()


def redefinir_senha(usuario, senha):
    salt = secrets.token_hex(16)
    senha_hash = hashlib.pbkdf2_hmac(
        'sha256', senha.encode('utf-8'), bytes.fromhex(salt), 600000
    ).hex()
    with conectar_banco() as conn:
        conn.execute(
            'UPDATE usuarios SET senha_hash = ?, salt = ? WHERE usuario = ?',
            (senha_hash, salt, usuario),
        )


def listar_usuarios_assinaturas():
    with conectar_banco() as conn:
        registros = conn.execute(
            '''SELECT usuarios.usuario, usuarios.nome, usuarios.nivel_acesso,
                      assinaturas_usuarios.atualizado_em
               FROM usuarios
               LEFT JOIN assinaturas_usuarios
                   ON assinaturas_usuarios.usuario = usuarios.usuario
               ORDER BY lower(usuarios.nome)'''
        ).fetchall()
    return [dict(registro) for registro in registros]


@st.cache_data
def carregar_assinatura_usuario(usuario):
    if not usuario:
        return None
    with conectar_banco() as conn:
        registro = conn.execute(
            'SELECT imagem_png FROM assinaturas_usuarios WHERE usuario = ?',
            (usuario,),
        ).fetchone()
    return bytes(registro[0]) if registro else None


def salvar_assinatura_usuario(usuario, imagem_png):
    agora = datetime.datetime.now(datetime.timezone.utc).isoformat()
    with conectar_banco() as conn:
        conn.execute(
            '''INSERT INTO assinaturas_usuarios (usuario, imagem_png, atualizado_em)
               VALUES (?, ?, ?)
               ON CONFLICT(usuario) DO UPDATE SET
                   imagem_png = excluded.imagem_png,
                   atualizado_em = excluded.atualizado_em''',
            (usuario, imagem_png, agora),
        )
    invalidar_cache_fichas()


def remover_assinatura_usuario(usuario):
    with conectar_banco() as conn:
        conn.execute(
            'DELETE FROM assinaturas_usuarios WHERE usuario = ?',
            (usuario,),
        )
    invalidar_cache_fichas()


def normalizar_imagem_assinatura(imagem):
    if len(imagem) > 5 * 1024 * 1024:
        raise ValueError("A imagem deve ter no máximo 5 MB.")

    try:
        with PillowImage.open(io.BytesIO(imagem)) as arquivo:
            if arquivo.format not in {"PNG", "JPEG"}:
                raise ValueError("Use uma imagem PNG ou JPEG.")
            if arquivo.width * arquivo.height > 20_000_000:
                raise ValueError("A imagem excede o limite de resolução permitido.")
            imagem_convertida = arquivo.convert("RGBA")
            imagem_convertida.thumbnail((1200, 400))
            saida = io.BytesIO()
            imagem_convertida.save(saida, format="PNG")
    except UnidentifiedImageError as erro:
        raise ValueError("O arquivo enviado não é uma imagem válida.") from erro
    except OSError as erro:
        raise ValueError("Não foi possível ler a imagem enviada.") from erro

    return saida.getvalue()


def registrar_ficha(prefixo, pesagem, revisao, usuario):
    agora = datetime.datetime.now(datetime.timezone.utc).isoformat()
    with conectar_banco() as conn:
        conn.execute(
            '''INSERT INTO fluxo_fichas
               (prefixo, pesagem, revisao, gerado_por, criado_em)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT(prefixo, pesagem, revisao) DO UPDATE SET
                   gerado_por = excluded.gerado_por,
                   aprovado_por = NULL,
                   criado_em = excluded.criado_em,
                   aprovado_em = NULL''',
            (
                safe_str(prefixo),
                normalizar_chave_ficha(pesagem),
                normalizar_chave_ficha(revisao),
                usuario,
                agora,
            )
        )
    invalidar_cache_fichas()

@st.cache_data
def buscar_fluxo_ficha(prefixo, pesagem, revisao):
    with conectar_banco() as conn:
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
            (
                safe_str(prefixo),
                normalizar_chave_ficha(pesagem),
                normalizar_chave_ficha(revisao),
            )
        ).fetchone()
    if not registro:
        return {
            'gerador_usuario': None,
            'gerador_nome': 'Não registrado',
            'aprovador_usuario': None,
            'aprovador_nome': 'Pendente',
            'aprovado_em': None,
        }
    return {
        'gerador_usuario': registro['gerador_usuario'],
        'gerador_nome': registro['gerador_nome'] or 'Não registrado',
        'aprovador_usuario': registro['aprovador_usuario'],
        'aprovador_nome': registro['aprovador_nome'] or 'Pendente',
        'aprovado_em': registro['aprovado_em'],
    }

def carregar_fluxos_fichas():
    with conectar_banco() as conn:
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
        (
            safe_str(registro['prefixo']),
            normalizar_chave_ficha(registro['pesagem']),
            normalizar_chave_ficha(registro['revisao']),
        ): {
            'gerador_usuario': registro['gerador_usuario'],
            'gerador_nome': registro['gerador_nome'] or 'Não registrado',
            'aprovador_usuario': registro['aprovador_usuario'],
            'aprovador_nome': registro['aprovador_nome'] or 'Pendente',
        }
        for registro in registros
    }


@st.cache_data
def carregar_campos_com_erro(prefixo, pesagem, revisao):
    with conectar_banco() as conn:
        registros = conn.execute(
            '''SELECT campos_com_erro.campo, campos_com_erro.descricao,
                      campos_com_erro.marcado_por, usuarios.nome AS marcador_nome
               FROM campos_com_erro
               LEFT JOIN usuarios ON usuarios.usuario = campos_com_erro.marcado_por
               WHERE prefixo = ? AND pesagem = ? AND revisao = ?
               ORDER BY campo''',
            (
                safe_str(prefixo),
                normalizar_chave_ficha(pesagem),
                normalizar_chave_ficha(revisao),
            ),
        ).fetchall()
    return [dict(registro) for registro in registros]


def carregar_chaves_fichas():
    with conectar_banco() as conn:
        registros = conn.execute(
            """SELECT dados->>'Prefixo' AS prefixo,
                      dados->>'Pesagem' AS pesagem,
                      COALESCE(dados->>'Revisao', dados->>'revisao') AS revisao
               FROM pesagens
               WHERE NULLIF(btrim(dados->>'Prefixo'), '') IS NOT NULL"""
        ).fetchall()
    return [dict(registro) for registro in registros]


@st.cache_data
def listar_fichas_devolvidas_usuario(usuario):
    with conectar_banco() as conn:
        registros = conn.execute(
            """SELECT campos_com_erro.prefixo, campos_com_erro.pesagem,
                      campos_com_erro.revisao, campos_com_erro.campo,
                      campos_com_erro.descricao,
                      fluxo_fichas.gerado_por AS gerador_usuario,
                      usuarios.nome AS gerador_nome
               FROM campos_com_erro
               JOIN fluxo_fichas
                 ON fluxo_fichas.prefixo = campos_com_erro.prefixo
                AND fluxo_fichas.pesagem = campos_com_erro.pesagem
                AND fluxo_fichas.revisao = campos_com_erro.revisao
               LEFT JOIN usuarios
                 ON usuarios.usuario = fluxo_fichas.gerado_por
               WHERE fluxo_fichas.gerado_por = ?
                 AND fluxo_fichas.aprovado_por IS NULL
               ORDER BY campos_com_erro.prefixo,
                        campos_com_erro.pesagem,
                        campos_com_erro.revisao,
                        campos_com_erro.campo""",
            (usuario,),
        ).fetchall()

    fichas = {}
    for registro in registros:
        chave = (
            safe_str(registro["prefixo"]),
            normalizar_chave_ficha(registro["pesagem"]),
            normalizar_chave_ficha(registro["revisao"]),
        )
        ficha = fichas.setdefault(
            chave,
            {
                "chave": chave,
                "pesagem_num": safe_float(chave[1]),
                "revisao_num": safe_float(chave[2]),
                "fluxo": {
                    "gerador_usuario": registro["gerador_usuario"],
                    "gerador_nome": registro["gerador_nome"] or "Não registrado",
                },
                "erros": [],
            },
        )
        ficha["erros"].append(
            {"campo": registro["campo"], "descricao": registro["descricao"]}
        )
    return list(fichas.values())


@st.cache_data
def listar_fichas_pendentes():
    fluxos = carregar_fluxos_fichas()
    with conectar_banco() as conn:
        registros_erros = conn.execute(
            '''SELECT prefixo, pesagem, revisao, campo, descricao
               FROM campos_com_erro
               ORDER BY prefixo, pesagem, revisao, campo'''
        ).fetchall()

    erros_por_ficha = {}
    for prefixo, pesagem, revisao, campo, descricao in registros_erros:
        chave = (
            safe_str(prefixo),
            normalizar_chave_ficha(pesagem),
            normalizar_chave_ficha(revisao),
        )
        erros_por_ficha.setdefault(chave, []).append(
            {"campo": campo, "descricao": descricao}
        )

    fichas = []
    ultima_revisao = {}
    for linha in carregar_chaves_fichas():
        prefixo = safe_str(linha.get("prefixo", ""))
        pesagem = normalizar_chave_ficha(linha.get("pesagem", ""))
        revisao = normalizar_chave_ficha(linha.get("revisao", ""))
        chave = (prefixo, pesagem, revisao)
        pesagem_num = safe_float(pesagem)
        revisao_num = safe_float(revisao)
        chave_pesagem = (prefixo, pesagem)
        ultima_revisao[chave_pesagem] = max(
            revisao_num,
            ultima_revisao.get(chave_pesagem, float("-inf")),
        )
        fichas.append(
            {
                "chave": chave,
                "chave_pesagem": chave_pesagem,
                "pesagem_num": pesagem_num,
                "revisao_num": revisao_num,
                "erros": erros_por_ficha.get(chave, []),
                "fluxo": fluxos.get(
                    chave,
                    {
                        "gerador_usuario": None,
                        "gerador_nome": "Não registrado",
                        "aprovador_usuario": None,
                        "aprovador_nome": "Pendente",
                    },
                ),
            }
        )

    pendentes_aprovacao = []
    pendentes_correcao = []
    for ficha in fichas:
        if ficha["erros"]:
            if not ficha["fluxo"]["aprovador_usuario"]:
                pendentes_correcao.append(ficha)
        elif (
            not ficha["fluxo"]["aprovador_usuario"]
            and ficha["revisao_num"]
            == ultima_revisao[ficha["chave_pesagem"]]
        ):
            pendentes_aprovacao.append(ficha)

    ordenar = lambda ficha: (
        ficha["chave"][0],
        ficha["pesagem_num"],
        ficha["revisao_num"],
    )
    return sorted(pendentes_aprovacao, key=ordenar), sorted(
        pendentes_correcao, key=ordenar
    )


def salvar_campos_com_erro(prefixo, pesagem, revisao, campos, usuario):
    agora = datetime.datetime.now(datetime.timezone.utc).isoformat()
    chave = (
        safe_str(prefixo),
        normalizar_chave_ficha(pesagem),
        normalizar_chave_ficha(revisao),
    )
    with conectar_banco() as conn:
        conn.execute(
            '''DELETE FROM campos_com_erro
               WHERE prefixo = ? AND pesagem = ? AND revisao = ?''',
            chave,
        )
        conn.executemany(
            '''INSERT INTO campos_com_erro
               (prefixo, pesagem, revisao, campo, descricao, marcado_por, marcado_em)
               VALUES (?, ?, ?, ?, ?, ?, ?)''',
            [
                (*chave, campo, descricao, usuario, agora)
                for campo, descricao in campos.items()
            ],
        )
    invalidar_cache_fichas()


def limpar_campos_com_erro(prefixo, pesagem, revisao):
    with conectar_banco() as conn:
        conn.execute(
            '''DELETE FROM campos_com_erro
               WHERE prefixo = ? AND pesagem = ? AND revisao = ?''',
            (
                safe_str(prefixo),
                normalizar_chave_ficha(pesagem),
                normalizar_chave_ficha(revisao),
            ),
        )
    invalidar_cache_fichas()

def aprovar_ficha(prefixo, pesagem, revisao, usuario):
    agora = datetime.datetime.now(datetime.timezone.utc).isoformat()
    with conectar_banco() as conn:
        conn.execute(
            '''INSERT INTO fluxo_fichas
               (prefixo, pesagem, revisao, aprovado_por, criado_em, aprovado_em)
               VALUES (?, ?, ?, ?, ?, ?)
               ON CONFLICT(prefixo, pesagem, revisao) DO UPDATE SET
                   aprovado_por = excluded.aprovado_por,
                   aprovado_em = excluded.aprovado_em''',
            (
                safe_str(prefixo),
                normalizar_chave_ficha(pesagem),
                normalizar_chave_ficha(revisao),
                usuario,
                agora,
                agora,
            )
        )
    invalidar_cache_fichas()

@st.cache_resource
def normalizar_chaves_fluxo():
    """Converte chaves antigas como '15.0' para '15' no fluxo de aprovação.

    O app procura as fichas por '15'; registros migrados com '15.0' ficavam
    invisíveis (gerador "Não registrado", devoluções que não apareciam).
    Linhas que colidiriam com uma já normalizada são mantidas como estão.
    """
    normalizar = r"regexp_replace({0}, '^(-?[0-9]+)[.]0+$', '\1')"
    for tabela, chave_extra in (("fluxo_fichas", ""), ("campos_com_erro", " AND g.campo = f.campo")):
        try:
            with conectar_banco() as conn:
                conn.execute(
                    f"""UPDATE {tabela} AS f
                        SET pesagem = {normalizar.format('f.pesagem')},
                            revisao = {normalizar.format('f.revisao')}
                        WHERE (f.pesagem ~ '[.]0+$' OR f.revisao ~ '[.]0+$')
                          AND NOT EXISTS (
                              SELECT 1 FROM {tabela} AS g
                              WHERE g.prefixo = f.prefixo
                                AND g.pesagem = {normalizar.format('f.pesagem')}
                                AND g.revisao = {normalizar.format('f.revisao')}
                                {chave_extra}
                          )"""
                )
        except Exception:
            pass


inicializar_controle_acesso()
normalizar_chaves_fluxo()
dict_tipos_aeronave.update(carregar_aeronaves_cadastradas())
df_historico = carregar_dados_banco()

# 3. FUNÇÃO DE EXPORTAÇÃO PARA EXCEL
def gerar_excel_por_template(dados, caminho_template="exemplo_ficha.xlsx"):
    try:
        wb = openpyxl.load_workbook(caminho_template)
    except FileNotFoundError as erro:
        raise FileNotFoundError(
            f"O arquivo de template '{caminho_template}' não foi encontrado na pasta."
        ) from erro

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

    deducoes = dados.get('deductions', [])
    if len(deducoes) > 22:
        raise ValueError("O template suporta no máximo 22 itens de deduções.")
    for idx in range(70, 92):
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

    adicoes = dados.get('additions', [])
    if len(adicoes) > 32:
        raise ValueError("O template suporta no máximo 32 itens de adições.")
    for idx in range(95, 127):
        if (idx - 95) < len(adicoes):
            a = adicoes[idx - 95]
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

    for celula, chave in (
        ("I50", "assinatura_emissor"),
        ("U50", "assinatura_aprovador"),
    ):
        assinatura = dados.get(chave)
        if assinatura:
            imagem = ExcelImage(io.BytesIO(assinatura))
            escala = min(180 / imagem.width, 54 / imagem.height, 1.5)
            imagem.width = int(imagem.width * escala * 1.5)
            imagem.height = int(imagem.height * escala)
            imagem.anchor = celula
            ws.add_image(imagem)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def mostrar_relatorio_pesagem(tabela, peso, braco, cg_mac):
    r1, r2, r3 = st.columns(3)
    r1.metric("Aircraft Basic Weight", f"{peso:,.2f} kg")
    r2.metric("Aircraft Basic Arm", f"{braco:,.4f} in")
    r3.metric("C.G. (% MAC)", f"{cg_mac:.2f} %")
    st.dataframe(
        pd.DataFrame(tabela),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Weight (kg)": st.column_config.NumberColumn(format="%.2f"),
            "Arm (inch)": st.column_config.NumberColumn(format="%.4f"),
            "Moment (kg x inch)": st.column_config.NumberColumn(format="%,.2f"),
        },
    )


# 4. TELAS E FUNCIONALIDADES

@protegido
def tela_consulta():
    cabecalho("Fichas", "Visão geral e consulta do histórico de pesagens.")
    if st.button("Atualizar", icon=":material/refresh:"):
        invalidar_cache_fichas()
        st.rerun()

    if df_historico.empty:
        st.warning(
            "Nenhuma ficha foi carregada do banco conectado ao aplicativo. "
            "No Streamlit Community Cloud, confira em Settings → Secrets se "
            "[connections.postgresql].url aponta para o mesmo banco Supabase "
            "usado na importação e reinicie o app."
        )
        return

    prefixos_unicos = sorted(df_historico['Prefixo'].dropna().unique().tolist()) if 'Prefixo' in df_historico.columns else []
    pendentes_aprovacao, pendentes_correcao = listar_fichas_pendentes()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Aeronaves com ficha", len(prefixos_unicos))
    k2.metric("Fichas emitidas", len(df_historico))
    k3.metric("Aguardando aprovação", len(pendentes_aprovacao))
    k4.metric("Devolvidas para correção", len(pendentes_correcao))

    ordenar_numero = lambda valores: sorted(valores, key=lambda v: (safe_float(v), v))
    with st.container(border=True):
        st.markdown("**Consultar ficha**")
        c1, c2, c3 = st.columns([2, 1, 1])
        prefixo = c1.selectbox(
            "Aeronave", prefixos_unicos, index=None, placeholder="Digite ou escolha o prefixo"
        )
        df_filtrado = (
            df_historico[df_historico['Prefixo'] == prefixo] if prefixo else df_historico.iloc[0:0]
        )
        pesagens = ordenar_numero(df_filtrado['Pesagem'].dropna().unique().tolist())
        pesagem = c2.selectbox(
            "Pesagem", pesagens, index=len(pesagens) - 1 if pesagens else None,
            disabled=not prefixo, placeholder="—",
        )
        revisoes = ordenar_numero(
            df_filtrado[df_filtrado['Pesagem'] == pesagem]['Revisao'].dropna().unique().tolist()
        ) if pesagem else []
        revisao = c3.selectbox(
            "Revisão", revisoes, index=len(revisoes) - 1 if revisoes else None,
            disabled=not pesagem, placeholder="—",
        )

    if prefixo and pesagem:
        if revisao:
                st.divider()
                linha = df_filtrado[(df_filtrado['Pesagem'] == pesagem) & (df_filtrado['Revisao'] == revisao)].iloc[0]
                renderizar_ficha_visualizacao(prefixo, pesagem, revisao, linha)

def renderizar_ficha_visualizacao(
    prefixo,
    pesagem,
    revisao,
    row,
    modo_aprovacao=False,
    destacar_erros=False,
):
    st.subheader(f"{prefixo} · Pesagem {pesagem} · Revisão {revisao}")
    fluxo = buscar_fluxo_ficha(prefixo, pesagem, revisao)
    st.caption(f"Gerada por: {fluxo['gerador_nome']} | Aprovada por: {fluxo['aprovador_nome']}")
    campos_pendentes = carregar_campos_com_erro(prefixo, pesagem, revisao)
    if campos_pendentes and not modo_aprovacao:
        st.warning(
            f"Ficha devolvida ao emissor ({fluxo['gerador_nome']}) para correção: "
            + "; ".join(campo['descricao'] for campo in campos_pendentes)
        )
    sinalizacoes_atuais = {
        campo['campo']: campo
        for campo in campos_pendentes
    }
    marcacoes_campos = {}
    rotulos_campos = {}

    def mostrar_campo(campo, rotulo, valor, container=st, permitir_erro=True):
        if not modo_aprovacao or not permitir_erro:
            if destacar_erros and campo in sinalizacoes_atuais:
                container.markdown(f":red[🔴 {rotulo} precisa de correção]")
            container.text_input(
                rotulo,
                value=safe_str(valor),
                disabled=True,
                key=(
                    f"visual_{'aprovacao' if modo_aprovacao else 'consulta'}_"
                    f"{prefixo}_{pesagem}_{revisao}_{campo}_"
                    f"{'erros' if destacar_erros else 'normal'}"
                ),
            )
            return

        coluna_valor, coluna_erro = container.columns([5, 1])
        coluna_valor.text_input(
            rotulo,
            value=safe_str(valor),
            disabled=True,
            key=f"visual_{prefixo}_{pesagem}_{revisao}_{campo}",
        )
        marcacoes_campos[campo] = coluna_erro.checkbox(
            "Erro",
            value=campo in sinalizacoes_atuais,
            key=f"aprovacao_erro_{prefixo}_{pesagem}_{revisao}_{campo}",
            help="Marque para solicitar ao emissor a correção deste campo.",
        )
        if marcacoes_campos[campo]:
            coluna_valor.markdown(f":red[🔴 {rotulo} marcado para correção]")
        rotulos_campos[campo] = f"{rotulo}: {safe_str(valor)}"

    def tabela_itens_aprovacao(tipo, campo_desc, campo_peso, campo_braco, limite):
        itens = []
        for indice in range(1, limite + 1):
            descricao = row.get(f'{tipo} {campo_desc} {indice}', None)
            if pd.isna(descricao) or not str(descricao).strip():
                continue
            peso = safe_float(row.get(f'{tipo} {campo_peso} {indice}', 0))
            braco = safe_float(row.get(f'{tipo} {campo_braco} {indice}', 0))
            itens.append((indice, safe_str(descricao), peso, braco))
        if not itens:
            st.caption("Nenhum item.")
            return
        larguras = [0.5, 4, 1.2, 1.2, 1.5, 0.8]
        cab = st.columns(larguras, vertical_alignment="center")
        for coluna, titulo in zip(cab, ["#", "Descrição", "Peso (kg)", "Braço (in)", "Momento", "Erro"]):
            coluna.markdown(f"<span style='color:#6B7280;font-size:.82rem;font-weight:600'>{titulo}</span>", unsafe_allow_html=True)
        for indice, descricao, peso, braco in itens:
            campo = f"{tipo} {campo_desc} {indice}"
            campos_item = {f"{tipo} {c} {indice}" for c in (campo_desc, campo_peso, campo_braco)}
            marcado_antes = bool(campos_item & set(sinalizacoes_atuais))
            st.markdown("<hr style='margin:.1rem 0;border:none;border-top:1px solid #EEF0F3'>", unsafe_allow_html=True)
            c = st.columns(larguras, vertical_alignment="center")
            c[0].write(str(indice))
            c[1].write(descricao)
            c[2].write(f"{peso:,.2f}")
            c[3].write(f"{braco:,.2f}")
            c[4].write(f"{peso * braco:,.2f}")
            marcado = c[5].checkbox(
                "Erro", value=marcado_antes, label_visibility="collapsed",
                key=f"aprovacao_erro_{prefixo}_{pesagem}_{revisao}_{campo}",
            )
            marcacoes_campos[campo] = marcado
            rotulos_campos[campo] = f"{tipo} item {indice}: {descricao}"
            if marcado:
                c[1].markdown(":red[🔴 marcado para correção]")

    aba1, aba2, aba3, aba4, aba5 = st.tabs(["Dados Gerais", "Células de Carga", "Deductions", "Additions", "Weighing Report"])
    info_aero = dict_tipos_aeronave.get(prefixo, {})
    lopa, config_lopa = separar_campos_lopa(row)

    with aba1:
        c_p, c_r = st.columns(2)
        mostrar_campo(
            "Pesagem", "Número da Pesagem:", row.get('Pesagem', ''),
            c_p, permitir_erro=False,
        )
        mostrar_campo(
            "Revisao", "Número da Revisão:", row.get('Revisao', ''),
            c_r, permitir_erro=False,
        )

        c1, c2, c3 = st.columns(3)
        mostrar_campo("Data da ficha", "Data de emissão:", row.get('Data da ficha', row.get('Data_da_Pesagem', '')), c1)
        mostrar_campo("Pesado Por", "Pesado por:", row.get('Pesado Por', row.get('WEIGHED BY', '')), c2)
        mostrar_campo("Local da pesagem", "Local da pesagem:", row.get('Local da pesagem', ''), c3)
        
        c4, c5, c6 = st.columns(3)
        mostrar_campo("Data_da_Pesagem", "Data da pesagem:", row.get('Data_da_Pesagem', ''), c4)
        mostrar_campo("lopa", "LOPA:", lopa, c5)
        mostrar_campo("config_lopa", "Configuração LOPA:", config_lopa, c6)
        
        c7, c8, c9 = st.columns(3)
        mostrar_campo("VRBL", "VRBL. NUMBER:", safe_str(row.get('VRBL')) or info_aero.get('vrbl', ''), c7)
        mostrar_campo("SERIAL", "SERIAL NUMBER:", safe_str(row.get('SERIAL')) or info_aero.get('serial', ''), c8)
        mostrar_campo("LINE", "LINE NUMBER:", safe_str(row.get('LINE')) or info_aero.get('line', ''), c9)
        mostrar_campo("Motivo", "Razão para emissão:", row.get('Motivo', ''))

    with aba2:
        campos_pesagem_1 = [
            ("Peso nariz LH", "Nariz LH"),
            ("Peso Nariz RH", "Nariz RH"),
            ("Peso MLG RH 1 ", "RH MLG 1"),
            ("Peso MLG LH 1", "LH MLG 1"),
            ("Peso MLG RH 2", "RH MLG 2"),
            ("Peso MLG LH2", "LH MLG 2"),
        ]
        campos_pesagem_2 = [
            ("Peso nariz LH pesagem 2", "Nariz LH"),
            ("Peso Nariz RH pesagem 2", "Nariz RH"),
            ("Peso MLG RH 1  pesagem 2", "RH MLG 1"),
            ("Peso MLG LH 1 pesagem 2", "LH MLG 1"),
            ("Peso MLG RH 2 pesagem 2", "RH MLG 2"),
            ("Peso MLG LH2 pesagem 2", "LH MLG 2"),
        ]
        for titulo, campos in (
            ("Pesagem 01", campos_pesagem_1),
            ("Pesagem 02", campos_pesagem_2),
        ):
            st.markdown(f"**{titulo}**")
            for indice in range(0, len(campos), 3):
                colunas = st.columns(3)
                for container, (campo, rotulo) in zip(
                    colunas, campos[indice:indice + 3]
                ):
                    alternativas = {
                        "Peso MLG RH 1 ": ["Peso MLG RH 1 ", "Peso MLG RH 1"],
                        "Peso MLG LH 1": ["Peso MLG LH 1", "Peso MLG LH 1 "],
                        "Peso MLG RH 2": ["Peso MLG RH 2", "Peso MLG RH 2 "],
                        "Peso MLG LH2": ["Peso MLG LH2", "Peso MLG LH 2"],
                        "Peso MLG RH 1  pesagem 2": [
                            "Peso MLG RH 1  pesagem 2",
                            "Peso MLG RH 1 pesagem 2",
                        ],
                    }
                    campo = get_real_col(alternativas.get(campo, [campo]))
                    valor = row.get(campo, "")
                    mostrar_campo(campo, rotulo, valor, container)

        st.markdown("**Medidas das canelas**")
        canela_lh, canela_rh = st.columns(2)
        mostrar_campo(
            "Canela MLG LH", "Canela MLG LH (in)",
            row.get("Canela MLG LH", ""), canela_lh,
        )
        mostrar_campo(
            "Canela MLG RH", "Canela MLG RH (in)",
            row.get("Canela MLG RH", ""), canela_rh,
        )

    with aba3:
        deducoes_lista = []
        for i in range(1, MAX_DEDUCTIONS + 1):
            desc = row.get(f'Deductions description {i}', None)
            peso = row.get(f'Deductions Weigth {i}', None)
            arm = row.get(f'Deductions arm {i}', None)
            if pd.notna(desc) and str(desc).strip() != '':
                deducoes_lista.append({"Descrição": desc, "Peso [Kg]": safe_float(peso), "Arm [pol]": safe_float(arm)})
        if modo_aprovacao:
            tabela_itens_aprovacao("Deductions", "description", "Weigth", "arm", MAX_DEDUCTIONS)
        elif deducoes_lista:
            st.dataframe(pd.DataFrame(deducoes_lista), use_container_width=True, hide_index=True)

    with aba4:
        adicoes_lista = []
        for i in range(1, MAX_ADDITIONS + 1):
            desc = row.get(f'Additions Description {i}', None)
            peso = row.get(f'Additions weigth {i}', None)
            arm = row.get(f'Additions arm {i}', None)
            if pd.notna(desc) and str(desc).strip() != '':
                adicoes_lista.append({"Descrição": desc, "Peso [Kg]": safe_float(peso), "Arm [pol]": safe_float(arm)})
        if modo_aprovacao:
            tabela_itens_aprovacao("Additions", "Description", "weigth", "arm", MAX_ADDITIONS)
        elif adicoes_lista:
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
        mostrar_relatorio_pesagem(report_data, basic_weight, basic_arm, cg_mac)
            
        pesagem_atual = safe_float(row.get('Pesagem', 0))
        revisao_atual = safe_float(row.get('Revisao', 0))
        
        df_aero = carregar_fichas_prefixo(prefixo)
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
            "vrbl_number": safe_str(safe_str(row.get('VRBL')) or info_aero.get('vrbl', '')),
            "serial_number": safe_str(safe_str(row.get('SERIAL')) or info_aero.get('serial', '')),
            "line_number": safe_str(safe_str(row.get('LINE')) or info_aero.get('line', '')),
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
        dados_excel["assinatura_emissor"] = carregar_assinatura_usuario(
            fluxo['gerador_usuario']
        )
        dados_excel["assinatura_aprovador"] = carregar_assinatura_usuario(
            fluxo['aprovador_usuario']
        )
        
        if modo_aprovacao:
            st.divider()
            st.subheader("Campos marcados para correção")
            campos_selecionados = {
                campo: rotulos_campos[campo]
                for campo, marcado in marcacoes_campos.items()
                if marcado
            }
            if campos_selecionados:
                for rotulo in campos_selecionados.values():
                    st.markdown(f"- {rotulo}")
                if st.button(
                    "Enviar para o emissor alterar",
                    type="secondary",
                    key=f"enviar_correcao_{prefixo}_{pesagem}_{revisao}",
                ):
                    salvar_campos_com_erro(
                        prefixo,
                        pesagem,
                        revisao,
                        campos_selecionados,
                        st.session_state['usuario_id'],
                    )
                    st.success(
                        "Os campos marcados foram enviados ao emissor para alteração."
                    )
                    st.rerun()
            elif campos_pendentes:
                st.info(
                    "Nenhum campo está marcado. Envie a lista vazia para remover "
                    "as sinalizações existentes."
                )
                if st.button(
                    "Limpar sinalizações",
                    key=f"limpar_correcao_{prefixo}_{pesagem}_{revisao}",
                ):
                    salvar_campos_com_erro(
                        prefixo,
                        pesagem,
                        revisao,
                        {},
                        st.session_state['usuario_id'],
                    )
                    st.success("As sinalizações foram removidas.")
                    st.rerun()
            else:
                st.caption("Marque “Erro” ao lado dos campos que precisam de correção.")

            pode_aprovar = (
                not campos_pendentes
                and not campos_selecionados
                and fluxo['gerador_usuario'] != st.session_state['usuario_id']
            )
            conferiu_cg = False
            if campos_pendentes:
                st.error(
                    "A aprovação está bloqueada até o emissor corrigir "
                    "os campos sinalizados."
                )
            elif fluxo['gerador_usuario'] == st.session_state['usuario_id']:
                st.warning("Quem emitiu a ficha não pode aprová-la.")
            else:
                conferiu_cg = st.checkbox(
                    "Confirmei os dados e o valor do C.G. (% MAC).",
                    key=f"conferiu_cg_{prefixo}_{pesagem}_{revisao}",
                )

            coluna_excel, coluna_aprovacao = st.columns(2)
            botoes_exportacao(
                coluna_excel, lambda: dados_excel, prefixo, pesagem, revisao,
                f"aprovacao_{prefixo}_{pesagem}_{revisao}",
            )
            if coluna_aprovacao.button(
                "Aprovar ficha",
                type="primary",
                disabled=not pode_aprovar or not conferiu_cg,
                key=f"aprovar_{prefixo}_{pesagem}_{revisao}",
                use_container_width=True,
            ):
                aprovar_ficha(
                    prefixo,
                    pesagem,
                    revisao,
                    st.session_state['usuario_id'],
                )
                st.rerun()
        else:
            botoes_exportacao(
                st, lambda: dados_excel, prefixo, pesagem, revisao,
                f"consulta_{prefixo}_{pesagem}_{revisao}_{destacar_erros}",
            )

def gerar_pdf_por_template(dados, caminho_template="exemplo_ficha.xlsx"):
    """Gera o Excel oficial e o converte em PDF com o LibreOffice."""
    import shutil
    import subprocess
    import tempfile

    executavel = shutil.which("soffice") or shutil.which("libreoffice")
    if not executavel:
        raise RuntimeError("LibreOffice não está instalado no servidor.")
    with tempfile.TemporaryDirectory() as pasta:
        caminho_xlsx = os.path.join(pasta, "ficha.xlsx")
        wb = openpyxl.load_workbook(
            io.BytesIO(gerar_excel_por_template(dados, caminho_template))
        )
        for planilha in wb.worksheets:
            # Ajusta a largura da ficha a uma página, como na impressão do Excel.
            planilha.sheet_properties.pageSetUpPr.fitToPage = True
            planilha.page_setup.fitToWidth = 1
            planilha.page_setup.fitToHeight = 0
            planilha.page_setup.paperSize = planilha.PAPERSIZE_A4
            # Só a ficha (colunas A:AD); à direita ficam tabelas auxiliares.
            planilha.print_area = f"A1:AD{min(planilha.max_row, 130)}"
        wb.save(caminho_xlsx)
        subprocess.run(
            [
                executavel, "--headless", "--norestore",
                f"-env:UserInstallation=file://{pasta}/perfil",
                "--convert-to", "pdf", "--outdir", pasta, caminho_xlsx,
            ],
            check=True, capture_output=True, timeout=120,
        )
        with open(os.path.join(pasta, "ficha.pdf"), "rb") as arquivo:
            return arquivo.read()


def nome_arquivo_ficha(prefixo, pesagem, revisao):
    return (
        f"{safe_str(prefixo)}_Pesagem_{normalizar_chave_ficha(pesagem)}"
        f"_Revisao_{normalizar_chave_ficha(revisao)}"
    )


def botoes_exportacao(container, obter_dados, prefixo, pesagem, revisao, chave):
    nome = nome_arquivo_ficha(prefixo, pesagem, revisao)
    col_excel, col_pdf = container.columns(2)
    col_excel.download_button(
        "Exportar Excel",
        icon=":material/table_view:",
        data=lambda: gerar_excel_por_template(obter_dados(), "exemplo_ficha.xlsx"),
        file_name=f"{nome}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key=f"excel_{chave}",
    )
    col_pdf.download_button(
        "Exportar PDF",
        icon=":material/picture_as_pdf:",
        data=lambda: gerar_pdf_por_template(obter_dados(), "exemplo_ficha.xlsx"),
        file_name=f"{nome}.pdf",
        mime="application/pdf",
        use_container_width=True,
        key=f"pdf_{chave}",
    )


# Função auxiliar para mapear as colunas corretas do Banco de Dados
def get_real_col(possible_names):
    colunas = st.session_state.get("historico_colunas", df_historico.columns)
    if len(colunas) == 0:
        return possible_names[0]
    for n in possible_names:
        if n in colunas:
            return n
    return possible_names[0]

MAX_DEDUCTIONS = 15
MAX_ADDITIONS = 16

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
        "Lavatory Dry Supplies(2EA)",
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

@st.cache_data
def carregar_presets_lopa():
    df = carregar_dados_banco()
    valores = set()
    for coluna in (" LOPA", "LOPA", "Configuração LOPA ", "Configuração LOPA"):
        if coluna in df.columns:
            valores.update(
                safe_str(valor)
                for valor in df[coluna].dropna().tolist()
                if safe_str(valor).casefold() not in {"lopa", "configuração lopa"}
            )
    lopa = sorted(
        {v for v in valores if v.upper().startswith("GLP-")} | {"GLP-MAX8-001-XMC"},
        key=str.casefold,
    )
    config = sorted(
        {v for v in valores if not v.upper().startswith("GLP-")} | {"186 Pax + 10 Flight Crew"},
        key=str.casefold,
    )
    return lopa, config


PRESETS["lopa"], PRESETS["config_lopa"] = carregar_presets_lopa()


@st.cache_data
def carregar_opcoes_cadastradas():
    with conectar_banco() as conn:
        registros = conn.execute(
            'SELECT categoria, valor FROM opcoes_cadastradas ORDER BY valor'
        ).fetchall()
    opcoes = {}
    for registro in registros:
        opcoes.setdefault(registro['categoria'], []).append(registro['valor'])
    return opcoes


def incluir_opcoes_preset(categoria, valores):
    existentes = {opcao.casefold() for opcao in PRESETS.setdefault(categoria, [])}
    for valor in valores:
        if valor.casefold() not in existentes:
            PRESETS[categoria].append(valor)
            existentes.add(valor.casefold())


for _categoria, _valores in carregar_opcoes_cadastradas().items():
    incluir_opcoes_preset(_categoria, _valores)


def salvar_opcao_digitada(categoria, chave_widget):
    """Grava no banco uma opção digitada que ainda não existe na lista."""
    valor = safe_str(st.session_state.get(chave_widget))
    if not valor or any(
        valor.casefold() == opcao.casefold() for opcao in PRESETS.get(categoria, [])
    ):
        return
    try:
        with conectar_banco() as conn:
            conn.execute(
                '''INSERT INTO opcoes_cadastradas
                   (categoria, valor, cadastrado_por, cadastrado_em)
                   VALUES (?, ?, ?, ?)
                   ON CONFLICT (categoria, valor) DO NOTHING''',
                (
                    categoria,
                    valor,
                    st.session_state.get('usuario_id'),
                    datetime.datetime.now(datetime.timezone.utc).isoformat(),
                ),
            )
    except Exception:
        st.toast(f"Não foi possível salvar “{valor}” na lista de opções.")
        return
    carregar_opcoes_cadastradas.clear()
    incluir_opcoes_preset(categoria, [valor])
    st.toast(f"“{valor}” foi adicionado à lista de opções.")


def campo_com_preset(container, label, valor, opcoes, key, categoria=None):
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
        on_change=salvar_opcao_digitada if categoria else None,
        args=(categoria, key) if categoria else None,
    )


def renderizar_editor_itens(
    container,
    itens,
    opcoes,
    key,
    momento_flaps=0,
    campos_com_erro=None,
    prefixo_campos_erro="",
    categoria=None,
):
    campos_com_erro = {safe_str(campo).casefold() for campo in campos_com_erro or ()}
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
        if prefixo_campos_erro and item:
            slot = item.get("slot", indice_item + 1)
            campos_item = {
                f"{prefixo_campos_erro} {campo} {slot}".casefold()
                for campo in ("description", "weigth", "arm")
            }
            if campos_com_erro & campos_item:
                descricao_col.markdown(":red[🔴 Este item foi devolvido para correção]")
        descricao = descricao_col.selectbox(
            f"Descrição, linha {posicao + 1}",
            escolhas,
            index=escolhas.index(descricao_inicial),
            key=f"{key}_{id_linha}_descricao",
            accept_new_options=True,
            placeholder="Digite ou selecione",
            label_visibility="collapsed",
            on_change=salvar_opcao_digitada if categoria else None,
            args=(categoria, f"{key}_{id_linha}_descricao") if categoria else None,
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

def formulario_pesagem(
    prefixo_selecionado,
    p_sugerida,
    r_sugerida,
    p_anterior,
    r_anterior,
    linha_existente=None,
    campos_com_erro=None,
):
    if linha_existente is None: linha_existente = {}
    campos_com_erro = set(campos_com_erro or ())

    def rotulo_form(campo, rotulo, container=st):
        if campo in campos_com_erro:
            container.markdown(f":red[🔴 Corrigir: {rotulo}]")
        return rotulo

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
        chave_data_emissao = get_real_col(['Data da ficha', 'Data_da_Pesagem'])
        chave_pesado_por = get_real_col(['Pesado Por', 'WEIGHED BY'])
        chave_local = get_real_col(['Local da pesagem'])
        data_emissao = c1.date_input(
            rotulo_form(chave_data_emissao, "Data de emissão:", c1),
            value=datetime.date.today(),
        )
        pesado_por = campo_com_preset(c2, rotulo_form(chave_pesado_por, "Pesado por:", c2), linha_existente.get(chave_pesado_por, ''), PRESETS["pesado_por"], f"{form_key}_pesado_por", categoria="pesado_por")
        local = campo_com_preset(c3, rotulo_form(chave_local, "Local da pesagem:", c3), linha_existente.get(chave_local, ''), PRESETS["local"], f"{form_key}_local", categoria="local")
        
        c4, c5, c6 = st.columns(3)
        key_data_pes = get_real_col(['Data_da_Pesagem', 'Data da ficha'])
        data_pesagem_existente = pd.to_datetime(
            linha_existente.get(key_data_pes), errors="coerce", dayfirst=True
        )
        val_pesagem = (
            data_pesagem_existente.date()
            if pd.notna(data_pesagem_existente)
            else datetime.date.today()
        )
        data_pesagem = c4.date_input(
            rotulo_form(get_real_col(['Data_da_Pesagem']), "Data da pesagem:", c4),
            value=val_pesagem,
        )
        
        lopa_inicial, config_lopa_inicial = separar_campos_lopa(linha_existente)
        lopa = campo_com_preset(c5, rotulo_form("lopa", "LOPA:", c5), lopa_inicial, PRESETS["lopa"], f"{form_key}_lopa")
        config_lopa = campo_com_preset(c6, rotulo_form("config_lopa", "Configuração LOPA:", c6), config_lopa_inicial, PRESETS["config_lopa"], f"{form_key}_config_lopa")
        
        c7, c8, c9 = st.columns(3)
        chave_vrbl = get_real_col(['VRBL', 'VRBL NUMBER'])
        chave_serial = get_real_col(['SERIAL', 'SERIAL NUMBER'])
        chave_line = get_real_col(['LINE', 'LINE NUMBER'])
        vrbl = c7.text_input(rotulo_form(chave_vrbl, "VRBL. NUMBER:", c7), value=safe_str(linha_existente.get(chave_vrbl)) or info_aero.get('vrbl', ''))
        serial = c8.text_input(rotulo_form(chave_serial, "SERIAL NUMBER:", c8), value=safe_str(linha_existente.get(chave_serial)) or info_aero.get('serial', ''))
        line = c9.text_input(rotulo_form(chave_line, "LINE NUMBER:", c9), value=safe_str(linha_existente.get(chave_line)) or info_aero.get('line', ''))

        chave_motivo = get_real_col(['Motivo', 'Razão'])
        razao = campo_com_preset(st, rotulo_form(chave_motivo, "Razão para emissão:"), linha_existente.get(chave_motivo, ''), PRESETS["motivo"], f"{form_key}_motivo", categoria="motivo")

    with aba2:
        st.markdown("**Pesagem 01**")
        p1_c1, p1_c2, p1_c3, p1_c4 = st.columns(4)
        chave_p1_nlh = get_real_col(['Peso nariz LH'])
        chave_p1_nrh = get_real_col(['Peso Nariz RH'])
        chave_p1_rhm = get_real_col(['Peso MLG RH 1 ', 'Peso MLG RH 1'])
        chave_p1_lhm = get_real_col(['Peso MLG LH 1', 'Peso MLG LH 1 '])
        p1_nlh = p1_c1.number_input(rotulo_form(chave_p1_nlh, "Nariz LH", p1_c1), value=safe_float(linha_existente.get(chave_p1_nlh, 0.0)), step=10.0)
        p1_nrh = p1_c2.number_input(rotulo_form(chave_p1_nrh, "Nariz RH", p1_c2), value=safe_float(linha_existente.get(chave_p1_nrh, 0.0)), step=10.0)
        p1_rhm = p1_c3.number_input(rotulo_form(chave_p1_rhm, "RH MLG 1", p1_c3), value=safe_float(linha_existente.get(chave_p1_rhm, 0.0)), step=10.0)
        p1_lhm = p1_c4.number_input(rotulo_form(chave_p1_lhm, "LH MLG 1", p1_c4), value=safe_float(linha_existente.get(chave_p1_lhm, 0.0)), step=10.0)
        
        p1_c5, p1_c6, p1_c7, p1_c8 = st.columns(4)
        chave_p1_rhm2 = get_real_col(['Peso MLG RH 2', 'Peso MLG RH 2 '])
        chave_p1_lhm2 = get_real_col(['Peso MLG LH2', 'Peso MLG LH 2'])
        p1_rhm2 = p1_c7.number_input(rotulo_form(chave_p1_rhm2, "RH MLG 2", p1_c7), value=safe_float(linha_existente.get(chave_p1_rhm2, 0.0)), step=10.0)
        p1_lhm2 = p1_c8.number_input(rotulo_form(chave_p1_lhm2, "LH MLG 2", p1_c8), value=safe_float(linha_existente.get(chave_p1_lhm2, 0.0)), step=10.0)

        st.markdown("**Pesagem 02**")
        p2_c1, p2_c2, p2_c3, p2_c4 = st.columns(4)
        chave_p2_nlh = get_real_col(['Peso nariz LH pesagem 2'])
        chave_p2_nrh = get_real_col(['Peso Nariz RH pesagem 2'])
        chave_p2_rhm = get_real_col(['Peso MLG RH 1  pesagem 2', 'Peso MLG RH 1 pesagem 2'])
        chave_p2_lhm = get_real_col(['Peso MLG LH 1 pesagem 2'])
        p2_nlh = p2_c1.number_input(rotulo_form(chave_p2_nlh, "Nariz LH P2", p2_c1), value=safe_float(linha_existente.get(chave_p2_nlh, 0.0)), step=10.0)
        p2_nrh = p2_c2.number_input(rotulo_form(chave_p2_nrh, "Nariz RH P2", p2_c2), value=safe_float(linha_existente.get(chave_p2_nrh, 0.0)), step=10.0)
        p2_rhm = p2_c3.number_input(rotulo_form(chave_p2_rhm, "RH MLG 1 P2", p2_c3), value=safe_float(linha_existente.get(chave_p2_rhm, 0.0)), step=10.0)
        p2_lhm = p2_c4.number_input(rotulo_form(chave_p2_lhm, "LH MLG 1 P2", p2_c4), value=safe_float(linha_existente.get(chave_p2_lhm, 0.0)), step=10.0)

        p2_c5, p2_c6, p2_c7, p2_c8 = st.columns(4)
        chave_p2_rhm2 = get_real_col(['Peso MLG RH 2 pesagem 2'])
        chave_p2_lhm2 = get_real_col(['Peso MLG LH2 pesagem 2'])
        p2_rhm2 = p2_c7.number_input(rotulo_form(chave_p2_rhm2, "RH MLG 2 P2", p2_c7), value=safe_float(linha_existente.get(chave_p2_rhm2, 0.0)), step=10.0)
        p2_lhm2 = p2_c8.number_input(rotulo_form(chave_p2_lhm2, "LH MLG 2 P2", p2_c8), value=safe_float(linha_existente.get(chave_p2_lhm2, 0.0)), step=10.0)
        
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
            rotulo_form("Canela MLG LH", f"Canela LH ({unidade_canela})", can_c1),
            step=0.1 if unidade_canela == "in" else 1.0,
            key=chave_canela_lh,
        )
        canela_rh_input = can_c2.number_input(
            rotulo_form("Canela MLG RH", f"Canela RH ({unidade_canela})", can_c2),
            step=0.1 if unidade_canela == "in" else 1.0,
            key=chave_canela_rh,
        )
        canela_lh = valor_canela_em_polegadas(canela_lh_input, unidade_canela)
        canela_rh = valor_canela_em_polegadas(canela_rh_input, unidade_canela)
        tail_val = 0.0

    with aba3:
        ded_ex = []
        for i in range(1, MAX_DEDUCTIONS + 1):
            desc = linha_existente.get(f'Deductions description {i}')
            if pd.notna(desc) and str(desc).strip():
                peso = safe_float(linha_existente.get(f'Deductions Weigth {i}'))
                arm = safe_float(linha_existente.get(f'Deductions arm {i}'))
                ded_ex.append({"Descrição": desc, "Peso (Kg)": peso, "Braço (in)": arm, "Momento (kg.in)": peso * arm, "slot": i})
                
        if not ded_ex:
            ded_ex = [{"Descrição": "Fuel (Usable)", "Peso (Kg)": 0.0, "Braço (in)": 660.5, "Momento (kg.in)": 0.0}]

        deducoes_editadas = renderizar_editor_itens(
            st,
            ded_ex,
            PRESETS["deductions"],
            f"{form_key}_deductions",
            campos_com_erro=campos_com_erro,
            prefixo_campos_erro="Deductions",
            categoria="deductions",
        )

    with aba4:
        momento_extra_flaps = momento_flaps_do_registro(linha_existente)
        if momento_extra_flaps is None:
            momento_extra_flaps = buscar_momento_flaps_banco(
                prefixo_selecionado, tipo_a
            )

        add_ex = []
        for i in range(1, MAX_ADDITIONS + 1):
            desc = linha_existente.get(f'Additions Description {i}')
            if pd.notna(desc) and str(desc).strip():
                peso = safe_float(linha_existente.get(f'Additions weigth {i}'))
                arm = safe_float(linha_existente.get(f'Additions arm {i}'))
                m_calc = peso * arm
                if "flap" in str(desc).casefold(): m_calc += momento_extra_flaps
                add_ex.append({"Descrição": desc, "Peso (Kg)": peso, "Braço (in)": arm, "Momento (kg.in)": m_calc, "slot": i})
                
        if not add_ex:
            add_ex = [{"Descrição": "Flaps 0 - 40° (up when weighed)", "Peso (Kg)": 0.0, "Braço (in)": 0.0, "Momento (kg.in)": momento_extra_flaps}]

        adicoes_editadas = renderizar_editor_itens(
            st,
            add_ex,
            PRESETS["additions"],
            f"{form_key}_additions",
            momento_flaps=momento_extra_flaps,
            campos_com_erro=campos_com_erro,
            prefixo_campos_erro="Additions",
            categoria="additions",
        )

    with aba5:
        angulo_salvo = normalizar_angulo_level_correction(
            linha_existente.get(get_real_col(['Graus correção do cg']), '')
        )
        opcoes_angulo = [""] + list(LEVEL_CORRECTION_VALUES)
        angulo_level_correction = st.selectbox(
            rotulo_form(
                get_real_col(['Graus correção do cg']),
                "Ângulo de inclinação",
            ),
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

        mostrar_relatorio_pesagem({
            "Reaction / Item": ["LH", "RH", "NOSE", "TAIL", "TOTAL REGISTERED", "LEVEL CORRECTION", "DEDUCTIONS", "ADDITIONS", "AIRCRAFT BASIC WEIGHT"],
            "Weight (kg)": [lh_val, rh_val, nose_val, tail_val, tot_reg_w, 0.0, ded_w, add_w, basic_w],
            "Arm (inch)": [arm_b_lh, arm_b_rh, arm_a, arm_c, tot_reg_arm, level_correction_factor, ded_arm, add_arm, basic_arm],
            "Moment (kg x inch)": [m_lh, m_rh, m_nose, m_tail, tot_reg_m, level_correction_moment, ded_m, add_m, basic_m]
        }, basic_w, basic_arm, cg_mac_val)
        
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
        dados_excel["assinatura_emissor"] = carregar_assinatura_usuario(
            st.session_state['usuario_id']
        )
        dados_excel["assinatura_aprovador"] = None
        
        st.session_state["dados_excel_temp"] = dados_excel

    # Mapeamento 100% seguro para evitar o Database Error
    novo_registro = linha_existente.copy()
    
    novo_registro[get_real_col(['Prefixo'])] = prefixo_selecionado
    if tipo_a:
        novo_registro[get_real_col(['Tipo_aeronave'])] = tipo_a
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
    
    # Todos os itens digitados são gravados. O filtro antigo por colunas do
    # histórico descartava, por exemplo, o 6º item e os campos VRBL/SERIAL/LINE.
    for i in range(1, MAX_DEDUCTIONS + 1):
        novo_registro[f'Deductions description {i}'] = None
        novo_registro[f'Deductions Weigth {i}'] = None
        novo_registro[f'Deductions arm {i}'] = None
    for i, row in enumerate(deducoes_editadas.to_dict('records')[:MAX_DEDUCTIONS], start=1):
        novo_registro[f'Deductions description {i}'] = row.get('Descrição')
        novo_registro[f'Deductions Weigth {i}'] = row.get('Peso (Kg)')
        novo_registro[f'Deductions arm {i}'] = row.get('Braço (in)')

    for i in range(1, MAX_ADDITIONS + 1):
        novo_registro[f'Additions Description {i}'] = None
        novo_registro[f'Additions weigth {i}'] = None
        novo_registro[f'Additions arm {i}'] = None
    for i, row in enumerate(adicoes_editadas.to_dict('records')[:MAX_ADDITIONS], start=1):
        novo_registro[f'Additions Description {i}'] = row.get('Descrição')
        novo_registro[f'Additions weigth {i}'] = row.get('Peso (Kg)')
        novo_registro[f'Additions arm {i}'] = row.get('Braço (in)')

    return novo_registro, id_pesagem_input, rev_input

@st.fragment
@protegido
def tela_nova_ficha():
    st.session_state["historico_colunas"] = ()

    if not exigir_nivel(NIVEL_OPERACAO, NIVEL_ADMIN):
        return
    cabecalho("Nova ficha", "Escolha a aeronave e preencha as abas em ordem. O Preview mostra o resultado antes de salvar.")
    ficha_salva = st.session_state.get("nova_ficha_salva")
    if ficha_salva:
        # Evita que um segundo clique em "Salvar" grave outra revisão igual.
        st.success(
            f"Ficha {ficha_salva[0]}, Pesagem {ficha_salva[1]}, "
            f"Revisão {ficha_salva[2]} salva e enviada para aprovação."
        )
        if st.button("Emitir outra ficha", type="primary"):
            st.session_state.pop("nova_ficha_salva", None)
            st.rerun()
        return

    usuario_atual = st.session_state['usuario_id']
    minhas_devolvidas = listar_fichas_devolvidas_usuario(usuario_atual)
    opcoes_correcao = {
        (
            f"{ficha['chave'][0]} | Pesagem {ficha['chave'][1]} | "
            f"Revisão {ficha['chave'][2]}"
        ): ficha
        for ficha in minhas_devolvidas
    }
    correcao_selecionada = None
    if opcoes_correcao:
        st.info("Há fichas devolvidas para correção. Selecione uma para revisar e reenviar.")
        rotulo_correcao = st.selectbox(
            "Ficha pendente de correção",
            ["Nova ficha"] + list(opcoes_correcao),
            key="ficha_devolvida_emissao",
        )
        correcao_selecionada = opcoes_correcao.get(rotulo_correcao)

    if correcao_selecionada:
        prefixo = correcao_selecionada["chave"][0]
    else:
        chaves_validas = [
            str(k)
            for k in dict_tipos_aeronave.keys()
            if str(k) not in ('nan', 'None', '')
        ]
        prefixo = st.selectbox(
            "Aeronave",
            [""] + sorted(chaves_validas),
            key="nova_ficha_aeronave",
        )

    if prefixo:
        st.divider()
        df_aero = carregar_fichas_prefixo(prefixo)
        st.session_state["historico_colunas"] = tuple(
            df_aero.columns if not df_aero.empty else df_historico.columns
        )
        if not df_aero.empty:
            df_aero['Pesagem_num'] = pd.to_numeric(
                df_aero['Pesagem'], errors='coerce'
            ).fillna(0)
            df_aero['Revisao_num'] = pd.to_numeric(
                df_aero['Revisao'], errors='coerce'
            ).fillna(0)

        if correcao_selecionada:
            ficha_correcao = df_aero.loc[
                (df_aero["Pesagem"].map(normalizar_chave_ficha)
                 == correcao_selecionada["chave"][1])
                & (df_aero["Revisao"].map(normalizar_chave_ficha)
                   == correcao_selecionada["chave"][2])
            ]
            if ficha_correcao.empty:
                st.error(
                    "A ficha devolvida não foi encontrada no histórico desta aeronave. "
                    "Atualize a lista e tente novamente."
                )
                return
            linha_base = ficha_correcao.iloc[0].to_dict()
            p_ant = int(correcao_selecionada["pesagem_num"])
            r_ant = int(correcao_selecionada["revisao_num"])
            p_nova = p_ant
            r_nova = r_ant
            campos_correcao = {
                campo["campo"] for campo in correcao_selecionada["erros"]
            }
            st.markdown("**Campos que precisam de correção:**")
            for erro in correcao_selecionada["erros"]:
                st.markdown(f":red[🔴 {erro['descricao']}]")
        elif not df_aero.empty:
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
            campos_correcao = set()
        else:
            st.info("🆕 **Nenhuma ficha anterior encontrada.** Iniciando a primeira Pesagem (1), Revisão (0) com a ficha em branco.")
            linha_base = {}
            p_nova = 1
            r_nova = 0
            p_ant = 0
            r_ant = 0
            campos_correcao = set()

        if not correcao_selecionada and p_ant > 0:
            campos_pendentes = carregar_campos_com_erro(prefixo, p_ant, r_ant)
            if campos_pendentes:
                emissor = buscar_fluxo_ficha(prefixo, p_ant, r_ant)['gerador_nome']
                st.warning(
                    f"A revisão anterior foi devolvida ao emissor ({emissor}). "
                    "Campos a corrigir: "
                    + "; ".join(campo['descricao'] for campo in campos_pendentes)
                )

        if not correcao_selecionada and st.checkbox(
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
            
        novo_registro, p_final, r_final = formulario_pesagem(
            prefixo,
            p_nova,
            r_nova,
            p_ant,
            r_ant,
            linha_existente=linha_base,
            campos_com_erro=campos_correcao,
        )
        
        st.divider()
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            label_salvar = (
                "Emitir ficha e enviar para aprovação"
                if correcao_selecionada
                else "Salvar Ficha"
            )
            if st.button(label_salvar, type="primary", use_container_width=True):
                ficha_existente = False
                if not df_aero.empty:
                    ficha_existente = (
                        (df_aero['Pesagem_num'] == int(p_final))
                        & (df_aero['Revisao_num'] == int(r_final))
                    ).any()

                novo_registro['Pesagem'] = str(int(p_final))
                novo_registro['Revisao'] = str(int(r_final))
                if correcao_selecionada:
                    if not atualizar_ficha(prefixo, p_final, r_final, novo_registro):
                        st.error(
                            "Não foi possível localizar a ficha original para "
                            "salvar a correção. Atualize a página e tente novamente."
                        )
                    else:
                        registrar_ficha(
                            prefixo,
                            p_final,
                            r_final,
                            st.session_state['usuario_id'],
                        )
                        limpar_campos_com_erro(prefixo, p_ant, r_ant)
                        st.session_state["nova_ficha_salva"] = (
                            prefixo, int(p_final), int(r_final)
                        )
                        st.rerun()
                elif ficha_existente:
                    st.error(f"Já existe uma ficha para {prefixo}, Pesagem {int(p_final)}, Revisão {int(r_final)}.")
                elif not inserir_ficha(novo_registro):
                    st.error(f"Já existe uma ficha para {prefixo}, Pesagem {int(p_final)}, Revisão {int(r_final)}.")
                else:
                    registrar_ficha(prefixo, p_final, r_final, st.session_state['usuario_id'])
                    st.session_state["nova_ficha_salva"] = (
                        prefixo, int(p_final), int(r_final)
                    )
                    st.rerun()
        with col_btn2:
            botoes_exportacao(
                st, lambda: st.session_state["dados_excel_temp"], prefixo,
                p_final, r_final, f"nova_{prefixo}",
            )

@st.fragment
@protegido
def tela_edicao():
    st.session_state["historico_colunas"] = ()
    cabecalho("Editar ficha", "Altere uma revisão já emitida. Ela volta para aprovação.")
    if not exigir_nivel(NIVEL_OPERACAO, NIVEL_ADMIN):
        return

    prefixos = [""] + carregar_prefixos_fichas()
    prefixo = st.selectbox("Aeronave", prefixos)
    
    if prefixo:
        df_filtrado = carregar_fichas_prefixo(prefixo)
        st.session_state["historico_colunas"] = tuple(df_filtrado.columns)
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
                r_nova = u_rev
                p_ant = u_pes
                r_ant = u_rev
                
                novo_registro, p_final, r_final = formulario_pesagem(prefixo, p_nova, r_nova, p_ant, r_ant, linha_existente=linha_atual)
                
                st.divider()
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    if st.button("Salvar alterações", type="primary", use_container_width=True):
                        fluxo_origem = buscar_fluxo_ficha(prefixo, pesagem, revisao)
                        novo_registro['Pesagem'] = str(int(p_final))
                        novo_registro['Revisao'] = str(int(r_final))
                        if not atualizar_ficha(
                            prefixo, pesagem, revisao, novo_registro
                        ):
                            st.error(
                                "Não foi possível localizar a ficha selecionada "
                                "para salvar as alterações. Atualize a página e tente novamente."
                            )
                            return
                        registrar_ficha(prefixo, p_final, r_final, st.session_state['usuario_id'])
                        if fluxo_origem['gerador_usuario'] == st.session_state['usuario_id']:
                            limpar_campos_com_erro(prefixo, pesagem, revisao)
                        st.success(
                            "Alterações salvas na mesma revisão e enviadas para aprovação."
                        )
                with col_btn2:
                    botoes_exportacao(
                        st, lambda: st.session_state["dados_excel_temp"], prefixo,
                        p_final, r_final, f"edicao_{prefixo}",
                    )

@protegido
def tela_criar_login():
    if not exigir_nivel(NIVEL_ADMIN):
        return

    cabecalho("Usuários", "Crie acessos para a equipe.")
    with st.form("form_criar_login"):
        nome = st.text_input("Nome completo")
        usuario = st.text_input("Usuário")
        senha = st.text_input("Senha", type="password")
        confirmar_senha = st.text_input("Confirmar senha", type="password")
        nivel = st.selectbox(
            "Perfil", list(PERFIS), format_func=PERFIS.get,
            help="Consulta só visualiza. Gera e aprova emite, edita e aprova fichas. "
                 "Administrador também gerencia usuários e assinaturas.",
        )
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
                criar_usuario(nome, usuario, senha, nivel)
                st.success(f"Login criado para {nome.strip()}.")
            except IntegrityError:
                st.error("Esse nome de usuário já está cadastrado.")

    st.divider()
    st.subheader("Alterar perfil")
    usuarios = listar_usuarios_assinaturas()
    if usuarios:
        opcoes = {u['usuario']: f"{u['nome']} ({u['usuario']})" for u in usuarios}
        perfis_atuais = {u['usuario']: u['nivel_acesso'] for u in usuarios}
        c1, c2, c3 = st.columns([2, 1.3, 1])
        escolhido = c1.selectbox("Usuário", list(opcoes), format_func=opcoes.get, key="perfil_usuario")
        novo_nivel = c2.selectbox(
            "Novo perfil", list(PERFIS), format_func=PERFIS.get,
            index=list(PERFIS).index(perfis_atuais.get(escolhido, NIVEL_CONSULTA))
            if perfis_atuais.get(escolhido) in PERFIS else 0,
            key=f"perfil_novo_{escolhido}",
        )
        c3.write("")
        c3.write("")
        if c3.button("Salvar perfil", use_container_width=True):
            if escolhido == st.session_state['usuario_id'] and novo_nivel != NIVEL_ADMIN:
                st.error("Você não pode remover o seu próprio acesso de Administrador.")
            else:
                with conectar_banco() as conn:
                    conn.execute(
                        'UPDATE usuarios SET nivel_acesso = ? WHERE usuario = ?',
                        (novo_nivel, escolhido),
                    )
                st.success(f"Perfil de {opcoes[escolhido]} alterado para {PERFIS[novo_nivel]}.")

        st.divider()
        st.subheader("Redefinir senha")
        st.caption("As senhas ficam criptografadas e não podem ser vistas; defina uma nova e informe à pessoa.")
        with st.form("form_redefinir_senha", clear_on_submit=True):
            alvo = st.selectbox("Usuário", list(opcoes), format_func=opcoes.get)
            c1, c2 = st.columns(2)
            nova = c1.text_input("Nova senha", type="password")
            confirmar = c2.text_input("Confirmar nova senha", type="password")
            trocar = st.form_submit_button("Redefinir senha", type="primary")
        if trocar:
            if not nova:
                st.error("Informe a nova senha.")
            elif nova != confirmar:
                st.error("As senhas não coincidem.")
            else:
                redefinir_senha(alvo, nova)
                st.success(f"Senha de {opcoes[alvo]} redefinida.")


@protegido
def tela_cadastrar_aeronave():
    if not exigir_nivel(NIVEL_OPERACAO, NIVEL_ADMIN):
        return

    cabecalho("Aeronaves", "Cadastre aeronaves que ainda não estão na frota do sistema.")
    st.caption(
        "A aeronave cadastrada aparece na lista de Nova Ficha. O modelo define "
        "o momento do flap sugerido na primeira ficha, a partir das fichas de "
        "outras aeronaves do mesmo modelo."
    )

    modelos = sorted({
        dados['modelo'] for dados in dict_tipos_aeronave.values()
        if dados.get('modelo')
    })
    modelo = st.selectbox(
        "Modelo da aeronave",
        [""] + modelos,
        key="cadastro_aeronave_modelo",
    )
    armnose_sugerido = armnose_padrao_modelo(modelo) if modelo else 93.0

    with st.form("form_cadastrar_aeronave", clear_on_submit=True):
        c1, c2 = st.columns(2)
        prefixo = c1.text_input("Prefixo", placeholder="PR-XXX")
        serial = c2.text_input("Serial number (MSN)")
        c3, c4, c5 = st.columns(3)
        vrbl = c3.text_input("VRBL number")
        line = c4.text_input("Line number")
        armnose = c5.number_input(
            "Arm nose (in)",
            value=armnose_sugerido,
            format="%.4f",
            help="Sugerido a partir das outras aeronaves do mesmo modelo.",
            key=f"cadastro_aeronave_armnose_{modelo}",
        )
        enviar = st.form_submit_button("Cadastrar aeronave", type="primary")

    if enviar:
        prefixo = normalizar_prefixo(prefixo)
        if not modelo:
            st.error("Selecione o modelo da aeronave.")
        elif not prefixo:
            st.error("Informe o prefixo.")
        elif any(normalizar_prefixo(k) == prefixo for k in dict_tipos_aeronave):
            st.error(f"A aeronave {prefixo} já está cadastrada.")
        else:
            try:
                cadastrar_aeronave(
                    prefixo,
                    modelo,
                    safe_identifier(serial),
                    safe_identifier(vrbl),
                    safe_identifier(line),
                    armnose,
                    st.session_state['usuario_id'],
                )
            except IntegrityError:
                st.error(f"A aeronave {prefixo} já está cadastrada.")
            else:
                st.session_state['aeronave_cadastrada'] = prefixo
                st.rerun()

    if st.session_state.get('aeronave_cadastrada'):
        st.success(
            f"Aeronave {st.session_state.pop('aeronave_cadastrada')} cadastrada."
        )

    cadastradas = carregar_aeronaves_cadastradas()
    st.divider()
    st.subheader("Aeronaves cadastradas pelo app")
    if not cadastradas:
        st.info("Nenhuma aeronave cadastrada pelo app ainda.")
        return

    st.dataframe(
        pd.DataFrame([
            {
                "Prefixo": prefixo,
                "Modelo": dados['modelo'],
                "Serial": dados['serial'],
                "VRBL": dados['vrbl'],
                "Line": dados['line'],
                "Arm nose": dados['armnose'],
                "Cadastrado por": dados['cadastrado_por'],
            }
            for prefixo, dados in cadastradas.items()
        ]),
        hide_index=True,
        use_container_width=True,
    )

    prefixos_com_ficha = set(carregar_prefixos_fichas())
    removiveis = [p for p in cadastradas if p not in prefixos_com_ficha]
    if removiveis:
        with st.expander("Remover aeronave cadastrada por engano"):
            st.caption("Só é possível remover aeronaves que ainda não têm ficha.")
            prefixo_remover = st.selectbox("Aeronave", removiveis)
            if st.button("Remover aeronave"):
                remover_aeronave_cadastrada(prefixo_remover)
                st.rerun()


@protegido
def tela_assinaturas():
    if not exigir_nivel(NIVEL_ADMIN):
        return

    cabecalho("Assinaturas", "Assinaturas usadas no Excel da ficha.")
    usuarios = listar_usuarios_assinaturas()
    if not usuarios:
        st.info("Não há usuários cadastrados.")
        return

    st.dataframe(
        pd.DataFrame([
            {
                "Nome": usuario['nome'],
                "Usuário": usuario['usuario'],
                "Perfil": PERFIS.get(usuario['nivel_acesso'], "—"),
                "Assinatura": "Cadastrada" if usuario['atualizado_em'] else "Pendente",
            }
            for usuario in usuarios
        ]),
        use_container_width=True,
        hide_index=True,
    )

    opcoes = {
        f"{usuario['nome']} ({usuario['usuario']})": usuario['usuario']
        for usuario in usuarios
    }
    usuario_selecionado = st.selectbox(
        "Selecione o usuário para cadastrar ou atualizar a assinatura",
        list(opcoes),
        key="usuario_assinatura",
    )
    usuario = opcoes[usuario_selecionado]
    assinatura_atual = carregar_assinatura_usuario(usuario)
    if assinatura_atual:
        st.image(assinatura_atual, caption="Assinatura cadastrada", width=300)

    arquivo = st.file_uploader(
        "Enviar assinatura (PNG ou JPEG, até 5 MB)",
        type=["png", "jpg", "jpeg"],
        key=f"arquivo_assinatura_{usuario}",
    )
    coluna_salvar, coluna_remover = st.columns(2)
    if coluna_salvar.button(
        "Salvar assinatura",
        type="primary",
        disabled=arquivo is None,
        key=f"salvar_assinatura_{usuario}",
        use_container_width=True,
    ):
        try:
            imagem_png = normalizar_imagem_assinatura(arquivo.getvalue())
        except ValueError as erro:
            st.error(str(erro))
        else:
            salvar_assinatura_usuario(usuario, imagem_png)
            st.success(f"Assinatura de {usuario_selecionado} salva.")
            st.rerun()

    if assinatura_atual and coluna_remover.button(
        "Remover assinatura",
        key=f"remover_assinatura_{usuario}",
        use_container_width=True,
    ):
        remover_assinatura_usuario(usuario)
        st.success(f"Assinatura de {usuario_selecionado} removida.")
        st.rerun()


@st.fragment
@protegido
def tela_aprovar_fichas():
    if not exigir_nivel(NIVEL_OPERACAO, NIVEL_ADMIN):
        return

    cabecalho("Aprovação", "Revise, devolva para correção ou aprove as fichas emitidas.")
    pendentes_aprovacao, pendentes_correcao = listar_fichas_pendentes()
    aba_aprovacao, aba_correcao = st.tabs(
        ["Pendentes de aprovação", "Pendentes de correção"]
    )

    with aba_aprovacao:
        if not pendentes_aprovacao:
            st.success("Não há fichas pendentes de aprovação.")
        else:
            opcoes_aprovacao = {
                (
                    f"{ficha['chave'][0]} | Pesagem {ficha['chave'][1]} | "
                    f"Revisão {ficha['chave'][2]}"
                ): ficha
                for ficha in pendentes_aprovacao
            }
            selecionada = st.selectbox(
                "Selecione a ficha para revisar",
                list(opcoes_aprovacao),
                key="ficha_pendente_aprovacao",
            )
            ficha = opcoes_aprovacao[selecionada]
            prefixo, pesagem, revisao = ficha["chave"]
            linha = carregar_linha_ficha(prefixo, pesagem, revisao)
            if linha is None:
                st.error("A ficha selecionada não foi encontrada no banco de dados.")
            else:
                renderizar_ficha_visualizacao(
                    prefixo,
                    pesagem,
                    revisao,
                    linha,
                    modo_aprovacao=True,
                )

    with aba_correcao:
        if not pendentes_correcao:
            st.success("Não há fichas aguardando correção do emissor.")
        else:
            opcoes_correcao = {
                (
                    f"{ficha['chave'][0]} | Pesagem {ficha['chave'][1]} | "
                    f"Revisão {ficha['chave'][2]}"
                ): ficha
                for ficha in pendentes_correcao
            }
            selecionada = st.selectbox(
                "Selecione a ficha devolvida",
                list(opcoes_correcao),
                key="ficha_pendente_correcao",
            )
            ficha = opcoes_correcao[selecionada]
            prefixo, pesagem, revisao = ficha["chave"]
            st.warning(
                f"Devolvida para correção por {ficha['fluxo']['gerador_nome']}. "
                "O emissor pode abrir “Nova Ficha”, selecionar esta ficha e "
                "reenviá-la para aprovação."
            )
            st.markdown("**Campos que precisam de correção:**")
            for erro in ficha["erros"]:
                st.markdown(f":red[🔴 {erro['descricao']}]")
            linha = carregar_linha_ficha(prefixo, pesagem, revisao)
            if linha is None:
                st.error("A ficha selecionada não foi encontrada no banco de dados.")
            else:
                renderizar_ficha_visualizacao(
                    prefixo,
                    pesagem,
                    revisao,
                    linha,
                    destacar_erros=True,
                )

def excluir_ficha(prefixo, pesagem, revisao):
    chave = (
        safe_str(prefixo),
        normalizar_chave_ficha(pesagem),
        normalizar_chave_ficha(revisao),
    )
    with conectar_banco() as conn:
        bloquear_prefixo(conn, prefixo)
        quantidade = 0
        for id_ficha in buscar_id_ficha(conn, prefixo, pesagem, revisao):
            quantidade += conn.execute(
                'DELETE FROM pesagens WHERE id = ?', (id_ficha,)
            ).rowcount
        if quantidade:
            conn.execute(
                '''DELETE FROM fluxo_fichas
                   WHERE prefixo = ? AND pesagem = ? AND revisao = ?''',
                chave,
            )
            conn.execute(
                '''DELETE FROM campos_com_erro
                   WHERE prefixo = ? AND pesagem = ? AND revisao = ?''',
                chave,
            )
    if quantidade:
        invalidar_cache_fichas()
    return quantidade


@protegido
def tela_excluir_ficha():
    if not exigir_nivel(NIVEL_OPERACAO, NIVEL_ADMIN):
        return

    cabecalho("Excluir ficha", "Remoção permanente de uma revisão.")
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
                st.success(f"Ficha excluída. Registros removidos: {removidas}.")
                st.rerun()
            else:
                st.error("A ficha não foi encontrada no banco de dados.")

# 5. ROTEAMENTO E BARRA LATERAL
def logo_html():
    for caminho in ("assets/logo_gol.svg", "assets/logo_gol.png"):
        if os.path.exists(caminho):
            import base64
            tipo = "svg+xml" if caminho.endswith(".svg") else "png"
            with open(caminho, "rb") as arquivo:
                dados = base64.b64encode(arquivo.read()).decode()
            return (
                f'<img src="data:image/{tipo};base64,{dados}" '
                'style="height:34px;width:auto;flex-shrink:0">'
            )
    return '<div class="marca-logo">W&B</div>'


MARCA_HTML = (
    f'<div class="marca">{logo_html()}<div>'
    '<div class="marca-nome">Pesagem e Balanceamento</div>'
    '<div class="marca-sub">Engenharia GOL</div></div></div>'
)

PAGINAS = {
    'consulta': ("Fichas", ":material/dashboard:", tela_consulta, {1, 2, 3}),
    'nova_ficha': ("Nova ficha", ":material/add_circle:", tela_nova_ficha, {1, 3}),
    'edicao': ("Editar ficha", ":material/edit_note:", tela_edicao, {1, 3}),
    'aprovar': ("Aprovação", ":material/task_alt:", tela_aprovar_fichas, {1, 3}),
    'cadastrar_aeronave': ("Aeronaves", ":material/flight:", tela_cadastrar_aeronave, {1, 3}),
    'criar_login': ("Usuários", ":material/group:", tela_criar_login, {3}),
    'assinaturas': ("Assinaturas", ":material/draw:", tela_assinaturas, {3}),
    'excluir': ("Excluir ficha", ":material/delete:", tela_excluir_ficha, {1, 3}),
}
GRUPOS_MENU = [
    ("Fichas", ['consulta', 'nova_ficha', 'edicao', 'aprovar']),
    ("Administração", ['cadastrar_aeronave', 'criar_login', 'assinaturas', 'excluir']),
]


def ir_para(pagina):
    st.session_state['pagina_atual'] = pagina


if not st.session_state['usuario_logado']:
    _, centro, _ = st.columns([1, 1.2, 1])
    with centro:
        st.markdown("<div style='height:12vh'></div>", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown(MARCA_HTML, unsafe_allow_html=True)
            st.caption("Entre com seu usuário para acessar as fichas.")
            with st.form("login_form", border=False):
                usuario = st.text_input("Usuário")
                senha = st.text_input("Senha", type="password")
                entrar = st.form_submit_button("Acessar", type="primary", use_container_width=True)
            if entrar:
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
    # Enquanto a pessoa digita a senha, deixa os dados das telas prontos.
    listar_fichas_pendentes()
    carregar_prefixos_fichas()
else:
    nivel = st.session_state['nivel_acesso']
    if nivel not in PAGINAS.get(st.session_state['pagina_atual'], (None, None, None, set()))[3]:
        st.session_state['pagina_atual'] = 'consulta'

    with st.sidebar:
        st.markdown(MARCA_HTML, unsafe_allow_html=True)
        perfil = PERFIS.get(nivel, "—")
        st.markdown(
            f'<div class="usuario">👤 <b>{st.session_state["nome_usuario"]}</b>'
            f'<br><span style="color:#6B7280">{perfil}</span></div>',
            unsafe_allow_html=True,
        )
        for grupo, paginas in GRUPOS_MENU:
            visiveis = [p for p in paginas if nivel in PAGINAS[p][3]]
            if not visiveis:
                continue
            st.markdown(f'<div class="grupo-menu">{grupo}</div>', unsafe_allow_html=True)
            for pagina in visiveis:
                rotulo, icone, _, _ = PAGINAS[pagina]
                st.button(
                    rotulo,
                    icon=icone,
                    key=f"menu_{pagina}",
                    use_container_width=True,
                    type="primary" if st.session_state['pagina_atual'] == pagina else "secondary",
                    on_click=ir_para,
                    args=(pagina,),
                )
        st.divider()
        if st.button("Sair", icon=":material/logout:", use_container_width=True):
            st.session_state['usuario_logado'] = False
            st.session_state['nivel_acesso'] = 0
            st.session_state['usuario_id'] = ""
            st.session_state['nome_usuario'] = ""
            st.session_state['pagina_atual'] = 'consulta'
            st.rerun()

    PAGINAS[st.session_state['pagina_atual']][2]()
