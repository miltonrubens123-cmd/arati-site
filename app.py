import os
import streamlit as st
import requests
import pandas as pd
from datetime import date
from calendar import monthrange

# ============================================================
# CONFIGURAÇÕES
# ============================================================

# Recomendado:
# 1) Revogue o token que foi exposto em print.
# 2) Cole o novo token aqui para teste local.
# 3) Depois podemos mover para .streamlit/secrets.toml.
SMARTSHEET_TOKEN = "Wmw5xFBAyj82NOIYU2yPObLR8FtRXm6c6QgV4"

SHEETS = {
    "contas": "6432342517540740",
    "fornecedores": "4516477865879428",
    "plano_contas": "8420259540559748",
}

BASE_URL = "https://api.smartsheet.com/2.0"

COL = {
    "cnpj": "CNPJ",
    "fornecedor": "Fornecedor",
    "fantasia": "Fantasia",
    "base": "Base",
    "entrada": "Entrada",
    "emissao": "Emissão",
    "vencimento": "Vencimento",
    "cod_plano": "Cod_Plano Contas",
    "descricao_conta": "Coluna1",
    "valor_original": "Valor Original",
    "valor_pago": "Valor Pago",
    "status": "Status",
    "data_pgto": "Data Pgto",
    "descricao": "Descrição",
    "parcela": "Parcela",
}

st.set_page_config(page_title="Contas a Pagar", layout="wide")

st.markdown(
    """
<style>
/* Reduz tamanho dos cards de métricas */
[data-testid="stMetric"] {
    background-color: #0F1623;
    border-radius: 10px;
    padding: 12px 14px;
}

/* Título do card */
[data-testid="stMetricLabel"] {
    font-size: 13px !important;
    white-space: normal !important;
    line-height: 1.2 !important;
}

/* Valor do card */
[data-testid="stMetricValue"] {
    font-size: 22px !important;
    line-height: 1.1 !important;
    white-space: normal !important;
    overflow-wrap: anywhere !important;
}

/* Container interno */
[data-testid="metric-container"] {
    overflow: visible !important;
}

/* Dá mais espaço horizontal no conteúdo principal */
.block-container {
    max-width: 1280px;
    padding-top: 1.5rem;
}
</style>
""",
    unsafe_allow_html=True,
)

USUARIOS = {
    "ARATI": {
        "admin": "Arati@1234",
    },
    "BREVES": {
        "admin": "Arati@1234",
    },
    "ITAITUBA": {
        "admin": "Arati@1234",
    },
    "SANTAREM": {
        "admin": "Arati@1234",
    },
    "ORIXIMINÁ": {
        "admin": "Arati@1234",
    },
    "NOVO PROGRESSO": {
        "admin": "Arati@1234",
    },
    "SALVATERRA": {
        "admin": "Arati@1234",
    },
}


def tela_login():
    st.markdown(
        """
    <style>
        .block-container {
            padding-top: 2rem;
            max-width: 1200px;
        }

        .login-card {
            background: #161B26;
            border: 1px solid #2A2F3A;
            border-radius: 18px;
            padding: 36px;
            box-shadow: 0 12px 30px rgba(0,0,0,.25);
        }

        .login-title {
            font-size: 32px;
            font-weight: 800;
            margin-bottom: 4px;
        }

        .login-subtitle {
            color: #AAB0BD;
            margin-bottom: 28px;
            font-size: 15px;
        }

        .brand-panel {
            background: linear-gradient(145deg, #E41E2B 0%, #9F111B 100%);
            border-radius: 22px;
            padding: 48px 36px;
            min-height: 520px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            box-shadow: 0 12px 30px rgba(0,0,0,.30);
        }

        .brand-title {
            font-size: 34px;
            font-weight: 900;
            color: white;
            line-height: 1.1;
            margin-top: 24px;
        }

        .brand-subtitle {
            color: rgba(255,255,255,.85);
            font-size: 16px;
            margin-top: 14px;
            max-width: 420px;
        }
    </style>
    """,
        unsafe_allow_html=True,
    )

    col_logo, col_login = st.columns([1.15, 1])

    with col_logo:

        st.image("app/imagens/logo.png", width=160)

        st.markdown("## Contas a pagar")

        st.markdown(
            """
            Controle interno de contas a pagar, lançamentos, 
            baixas e acompanhamento financeiro.

            Integrado com Smartsheet para gestão ágil e eficiente das finanças.
            """
        )

        st.markdown("</div>", unsafe_allow_html=True)

    with col_login:
        st.markdown("<div class='login-card'>", unsafe_allow_html=True)

        st.markdown(
            """
            <div class='login-title'>Acesso ao sistema</div>
            <div class='login-subtitle'>Informe filial, usuário e senha para continuar.</div>
        """,
            unsafe_allow_html=True,
        )

        with st.form("login"):
            filial = st.selectbox("Filial", list(USUARIOS.keys()))
            usuario = st.text_input("Usuário")
            senha = st.text_input("Senha", type="password")

            entrar = st.form_submit_button(
                "Entrar", type="primary", use_container_width=True
            )

        st.markdown("</div>", unsafe_allow_html=True)

    if entrar:
        usuario_limpo = usuario.strip().lower()

        if (
            usuario_limpo in USUARIOS[filial]
            and USUARIOS[filial][usuario_limpo] == senha
        ):
            st.session_state["autenticado"] = True
            st.session_state["filial"] = filial
            st.session_state["usuario"] = usuario_limpo
            st.rerun()
        else:
            st.error("Filial, usuário ou senha inválidos.")


if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    tela_login()
    st.stop()
# ============================================================
# API SMARTSHEET
# ============================================================


def headers_json():
    return {
        "Authorization": f"Bearer {SMARTSHEET_TOKEN}",
        "Content-Type": "application/json",
    }


def headers_auth():
    return {"Authorization": f"Bearer {SMARTSHEET_TOKEN}"}


def validar_token():
    if not SMARTSHEET_TOKEN or SMARTSHEET_TOKEN == "COLE_SEU_TOKEN_NOVO_AQUI":
        st.error("Configure o SMARTSHEET_TOKEN no início do arquivo app.py.")
        st.stop()


def api_request(method, path, **kwargs):
    validar_token()

    url = f"{BASE_URL}{path}"
    try:
        r = requests.request(method, url, timeout=60, **kwargs)
    except requests.RequestException as e:
        st.error(f"Erro de conexão com Smartsheet: {e}")
        st.stop()

    if r.status_code not in [200, 201]:
        st.error(f"Erro na API Smartsheet ({r.status_code}): {r.text}")
        st.stop()

    return r.json() if r.text else {}


@st.cache_data(ttl=600, show_spinner=False)
def carregar_sheet(sheet_id):
    # exclude=attachments melhora tempo de resposta quando há muitos anexos.
    return api_request(
        "GET",
        f"/sheets/{sheet_id}?exclude=attachments,discussions,filters",
        headers=headers_json(),
    )


def atualizar_linhas_contas(payload):
    return api_request(
        "PUT",
        f"/sheets/{SHEETS['contas']}/rows",
        headers=headers_json(),
        json=payload,
    )


def criar_linhas_contas(payload):
    return api_request(
        "POST",
        f"/sheets/{SHEETS['contas']}/rows",
        headers=headers_json(),
        json=payload,
    )


def upload_anexo(row_id, arquivo):
    validar_token()

    url = f"{BASE_URL}/sheets/{SHEETS['contas']}/rows/{row_id}/attachments"
    files = {"file": (arquivo.name, arquivo.getvalue())}

    try:
        r = requests.post(url, headers=headers_auth(), files=files, timeout=120)
    except requests.RequestException as e:
        st.error(f"Erro de conexão ao anexar arquivo: {e}")
        st.stop()

    if r.status_code not in [200, 201]:
        st.error(f"Erro ao anexar arquivo: {r.text}")
        st.stop()

    return r.json()


# ============================================================
# TRANSFORMAÇÃO / PERFORMANCE
# ============================================================


def mapa_colunas(sheet):
    return {c["title"]: c["id"] for c in sheet.get("columns", [])}


def sheet_para_dataframe(sheet):
    col_id_para_nome = {c["id"]: c["title"] for c in sheet.get("columns", [])}

    linhas = []
    for row in sheet.get("rows", []):
        item = {"row_id": row["id"]}
        for cell in row.get("cells", []):
            titulo = col_id_para_nome.get(cell.get("columnId"))
            if titulo:
                item[titulo] = cell.get("displayValue", cell.get("value"))
        linhas.append(item)

    return pd.DataFrame(linhas), mapa_colunas(sheet)


def valor_numero_serie(s):
    if s is None:
        return pd.Series(dtype=float)

    return (
        s.fillna("")
        .astype(str)
        .str.replace("R$", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.strip()
        .replace("", "0")
        .pipe(pd.to_numeric, errors="coerce")
        .fillna(0.0)
    )


def preparar_contas(df):
    if df.empty:
        return df

    df = df.copy()

    if COL["valor_original"] in df.columns:
        df["Valor Original Num"] = valor_numero_serie(df[COL["valor_original"]])
    else:
        df["Valor Original Num"] = 0.0

    if COL["valor_pago"] in df.columns:
        df["Valor Pago Num"] = valor_numero_serie(df[COL["valor_pago"]])
    else:
        df["Valor Pago Num"] = 0.0

    if COL["vencimento"] in df.columns:
        df["Data Vencimento Ajustada"] = pd.to_datetime(
            df[COL["vencimento"]], errors="coerce"
        ).dt.date
    else:
        df["Data Vencimento Ajustada"] = pd.NaT

    if COL["data_pgto"] in df.columns:
        df["Data Pgto Ajustada"] = pd.to_datetime(
            df[COL["data_pgto"]], errors="coerce"
        ).dt.date
    else:
        df["Data Pgto Ajustada"] = pd.NaT

    if COL["descricao_conta"] in df.columns:
        df["Descrição conta"] = df[COL["descricao_conta"]].fillna("").astype(str)
    else:
        df["Descrição conta"] = ""

    if COL["status"] in df.columns:
        df["Status Normalizado"] = (
            df[COL["status"]].fillna("").astype(str).str.upper().str.strip()
        )
    else:
        df["Status Normalizado"] = ""

    return df


def texto_pesquisa_df(df):
    if df.empty:
        return pd.Series(dtype=str)

    return df.fillna("").astype(str).agg(" | ".join, axis=1).str.upper()


@st.cache_data(ttl=600, show_spinner=False)
def carregar_bases_processadas():
    sheet_contas = carregar_sheet(SHEETS["contas"])
    sheet_fornecedores = carregar_sheet(SHEETS["fornecedores"])
    sheet_plano = carregar_sheet(SHEETS["plano_contas"])

    df_contas, col_map_contas = sheet_para_dataframe(sheet_contas)
    df_forn, _ = sheet_para_dataframe(sheet_fornecedores)
    df_plano, _ = sheet_para_dataframe(sheet_plano)

    df_contas = preparar_contas(df_contas)

    return df_contas, df_forn, df_plano, col_map_contas


def formatar_moeda(v):
    return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def limitar_opcoes(lista, limite=300):
    return lista[:limite]


def apenas_numeros(valor):
    return "".join(filter(str.isdigit, str(valor)))


def primeira_coluna_existente(df, nomes_possiveis):
    if df.empty:
        return None

    mapa = {str(c).strip().upper(): c for c in df.columns}

    for nome in nomes_possiveis:
        chave = str(nome).strip().upper()
        if chave in mapa:
            return mapa[chave]

    # fallback por contém
    for nome in nomes_possiveis:
        alvo = str(nome).strip().upper()
        for col in df.columns:
            if alvo in str(col).strip().upper():
                return col

    return None


def colunas_fornecedores(df_forn):
    return {
        "cnpj": primeira_coluna_existente(
            df_forn,
            [
                COL["cnpj"],
                "CNPJ",
                "CNPJ/CPF",
                "CPF/CNPJ",
                "CPF CNPJ",
                "CNPJ_CPF",
                "CGC",
                "CPF",
            ],
        ),
        "fornecedor": primeira_coluna_existente(
            df_forn,
            [
                COL["fornecedor"],
                "FORNECEDOR",
                "RAZÃO SOCIAL",
                "RAZAO SOCIAL",
                "NOME",
                "NOME FORNECEDOR",
                "RAZ_SOCIAL",
            ],
        ),
        "fantasia": primeira_coluna_existente(
            df_forn,
            [
                COL["fantasia"],
                "FANTASIA",
                "NOME FANTASIA",
                "APELIDO",
            ],
        ),
    }


def colunas_plano(df_plano):
    return {
        "cod": primeira_coluna_existente(
            df_plano,
            [
                COL["cod_plano"],
                "COD_PLANO CONTAS",
                "COD PLANO CONTAS",
                "COD_PLANO",
                "CODIGO",
                "CÓDIGO",
                "CODIGO PLANO",
                "CÓDIGO PLANO",
                "CONTA",
                "COD_CONTA",
            ],
        ),
        "desc": primeira_coluna_existente(
            df_plano,
            [
                COL["descricao_conta"],
                "COLUNA1",
                "DESCRIÇÃO",
                "DESCRICAO",
                "DESCRIÇÃO CONTA",
                "DESCRICAO CONTA",
                "CONTA DESCRIÇÃO",
                "CONTA DESCRICAO",
            ],
        ),
    }


def montar_opcoes_fornecedores(df_forn, busca=""):
    if df_forn.empty:
        return []

    cols = colunas_fornecedores(df_forn)
    cnpj_col = cols["cnpj"]
    fornecedor_col = cols["fornecedor"]
    fantasia_col = cols["fantasia"]

    if not cnpj_col:
        return []

    df = df_forn.copy()

    cnpj_original = (
        df.get(cnpj_col, pd.Series("", index=df.index))
        .fillna("")
        .astype(str)
        .str.strip()
    )
    cnpj_limpo = cnpj_original.apply(apenas_numeros)

    fornecedor = (
        df.get(fornecedor_col, pd.Series("", index=df.index))
        .fillna("")
        .astype(str)
        .str.strip()
        if fornecedor_col
        else pd.Series("", index=df.index)
    )

    fantasia = (
        df.get(fantasia_col, pd.Series("", index=df.index))
        .fillna("")
        .astype(str)
        .str.strip()
        if fantasia_col
        else pd.Series("", index=df.index)
    )

    df_opcoes = pd.DataFrame(
        {
            "cnpj_limpo": cnpj_limpo,
            "cnpj_original": cnpj_original,
            "fornecedor": fornecedor,
            "fantasia": fantasia,
        }
    )

    termo = str(busca or "").strip().upper()
    termo_num = apenas_numeros(termo)

    if termo:
        texto = (
            df_opcoes["cnpj_limpo"]
            + " | "
            + df_opcoes["cnpj_original"]
            + " | "
            + df_opcoes["fornecedor"]
            + " | "
            + df_opcoes["fantasia"]
        ).str.upper()

        if termo_num:
            mask = texto.str.contains(termo, na=False, regex=False) | df_opcoes[
                "cnpj_limpo"
            ].str.contains(termo_num, na=False, regex=False)
        else:
            mask = texto.str.contains(termo, na=False, regex=False)

        df_opcoes = df_opcoes[mask]

    df_opcoes["label"] = (
        df_opcoes["cnpj_limpo"]
        + " | "
        + df_opcoes["cnpj_original"]
        + " | "
        + df_opcoes["fornecedor"]
        + " | "
        + df_opcoes["fantasia"]
    )

    return df_opcoes["label"].drop_duplicates().head(500).tolist()


def montar_opcoes_plano(df_plano, busca=""):
    if df_plano.empty:
        return []

    # força encontrar as colunas
    colunas = [c.strip().upper() for c in df_plano.columns]

    cod_col = None
    desc_col = None

    for c in df_plano.columns:
        nome = c.strip().upper()

        if "COD" in nome or "CÓD" in nome:
            cod_col = c

        if "DESC" in nome or "DESCRI" in nome or "CONTA" in nome:
            desc_col = c

    if not cod_col:
        return []

    df = df_plano.copy()

    cod = df[cod_col].fillna("").astype(str).str.strip()
    desc = df[desc_col].fillna("").astype(str).str.strip() if desc_col else ""

    df_opcoes = pd.DataFrame({"cod": cod, "desc": desc})

    # filtro de busca
    termo = str(busca or "").upper().strip()
    if termo:
        mask = df_opcoes["cod"].str.upper().str.contains(termo, na=False) | df_opcoes[
            "desc"
        ].str.upper().str.contains(termo, na=False)
        df_opcoes = df_opcoes[mask]

    # 👉 AQUI É O PRINCIPAL
    df_opcoes["label"] = df_opcoes["cod"] + " | " + df_opcoes["desc"]

    return df_opcoes["label"].drop_duplicates().tolist()


def obter_fornecedor_por_cnpj(df_forn, cnpj_final):
    if df_forn.empty or not cnpj_final:
        return ""

    cols = colunas_fornecedores(df_forn)
    cnpj_col = cols["cnpj"]
    fornecedor_col = cols["fornecedor"]
    fantasia_col = cols["fantasia"]

    if not cnpj_col:
        return ""

    alvo = apenas_numeros(cnpj_final)

    df = df_forn.copy()
    df["_cnpj_limpo"] = df[cnpj_col].fillna("").astype(str).apply(apenas_numeros)

    linha = df[df["_cnpj_limpo"] == alvo]

    if linha.empty:
        return ""

    row = linha.iloc[0]

    if fornecedor_col and str(row.get(fornecedor_col, "")).strip():
        return str(row.get(fornecedor_col, "")).strip()

    if fantasia_col:
        return str(row.get(fantasia_col, "")).strip()

    return ""


def obter_descricao_plano(df_plano, cod_plano):
    if df_plano.empty or not cod_plano:
        return ""

    cols = colunas_plano(df_plano)
    cod_col = cols["cod"]
    desc_col = cols["desc"]

    if not cod_col or not desc_col:
        return ""

    linha = df_plano[
        df_plano[cod_col].fillna("").astype(str).str.strip() == str(cod_plano).strip()
    ]

    if linha.empty:
        return ""

    return str(linha.iloc[0].get(desc_col, "")).strip()


# ============================================================
# AÇÕES
# ============================================================


def criar_lancamentos(col_map, dados_parcelas):
    rows = []

    for item in dados_parcelas:
        cells = []

        for campo, valor in item.items():
            nome_coluna = COL.get(campo)

            if nome_coluna and nome_coluna in col_map:
                cells.append(
                    {
                        "columnId": col_map[nome_coluna],
                        "value": valor,
                        "strict": False,
                    }
                )

        rows.append({"toBottom": True, "cells": cells})

    return criar_linhas_contas(rows)


def marcar_pago(row_id, col_map, valor_pago, data_pgto, anotar_saldo_zero=False):
    cells = [
        {
            "columnId": col_map[COL["status"]],
            "value": "Pago",
            "strict": False,
        },
        {
            "columnId": col_map[COL["valor_pago"]],
            "value": float(valor_pago),
            "strict": False,
        },
        {
            "columnId": col_map[COL["data_pgto"]],
            "value": str(data_pgto),
            "strict": False,
        },
    ]

    # Só ativa se sua sheet tiver coluna Saldo e você quiser gravar 0.
    if anotar_saldo_zero and "Saldo" in col_map:
        cells.append(
            {
                "columnId": col_map["Saldo"],
                "value": 0,
                "strict": False,
            }
        )

    payload = [{"id": int(row_id), "cells": cells}]

    return atualizar_linhas_contas(payload)


# ============================================================
# INTERFACE
# ============================================================

st.markdown("### 📊 Visão Geral")
st.divider()

with st.sidebar:
    st.caption(f"Filial: {st.session_state.get('filial', '-')}")
    st.caption(f"Usuário: {st.session_state.get('usuario', '-')}")

    pagina = st.radio("Menu", ["Dashboard", "Lançamentos e Baixa"])

    if st.button("Atualizar dados"):
        st.cache_data.clear()
        st.rerun()

    if st.button("Sair"):
        st.session_state.clear()
        st.rerun()

with st.spinner("Carregando dados do Smartsheet..."):
    df_contas, df_forn, df_plano, col_map_contas = carregar_bases_processadas()

# ========================
# BASES DISPONÍVEIS
# ========================
bases_opcoes = []

if not df_contas.empty and COL["base"] in df_contas.columns:
    bases_opcoes = (
        df_contas[COL["base"]]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda x: x != ""]
        .unique()
        .tolist()
    )

bases_opcoes = sorted(bases_opcoes)


# ============================================================
# DASHBOARD
# ============================================================

if pagina == "Dashboard":
    hoje = date.today()

    col1, col2, col3, col4 = st.columns(4)

    ano = col1.selectbox(
        "Ano",
        list(range(hoje.year - 5, hoje.year + 3)),
        index=5,
    )

    mes = col2.selectbox(
        "Mês",
        list(range(1, 13)),
        index=hoje.month - 1,
    )

    if df_contas.empty:
        st.warning("Nenhum lançamento encontrado.")
        st.stop()

    ini_mes = date(ano, mes, 1)
    fim_mes = date(ano, mes, monthrange(ano, mes)[1])

    df_mes = df_contas[
        (
            (df_contas["Data Vencimento Ajustada"].notna())
            & (df_contas["Data Vencimento Ajustada"] >= ini_mes)
            & (df_contas["Data Vencimento Ajustada"] <= fim_mes)
        )
        | (
            (df_contas["Data Pgto Ajustada"].notna())
            & (df_contas["Data Pgto Ajustada"] >= ini_mes)
            & (df_contas["Data Pgto Ajustada"] <= fim_mes)
        )
    ].copy()

    descricoes = ["Todas"]
    if not df_mes.empty:
        descricoes += sorted(
            df_mes["Descrição conta"]
            .dropna()
            .astype(str)
            .loc[lambda x: x.str.strip() != ""]
            .unique()
            .tolist()
        )

    filtro_descricao = col3.selectbox(
        "Descrição conta",
        descricoes,
        index=0,
        placeholder="Pesquise a descrição",
    )

    status_opcoes = ["Todos"]
    if COL["status"] in df_mes.columns and not df_mes.empty:
        status_opcoes += sorted(
            df_mes[COL["status"]]
            .dropna()
            .astype(str)
            .loc[lambda x: x.str.strip() != ""]
            .unique()
            .tolist()
        )

    filtro_status = col4.selectbox(
        "Status",
        status_opcoes,
        index=0,
        placeholder="Pesquise o status",
    )

    if filtro_descricao != "Todas":
        df_mes = df_mes[df_mes["Descrição conta"].astype(str) == filtro_descricao]

    if filtro_status != "Todos" and COL["status"] in df_mes.columns:
        df_mes = df_mes[df_mes[COL["status"]].astype(str) == filtro_status]

    df_a_pagar = df_mes[
        (df_mes["Data Vencimento Ajustada"].notna())
        & (df_mes["Data Vencimento Ajustada"] >= ini_mes)
        & (df_mes["Data Vencimento Ajustada"] <= fim_mes)
        & (df_mes["Status Normalizado"] != "PAGO")
    ]

    df_baixado = df_mes[
        (df_mes["Data Pgto Ajustada"].notna())
        & (df_mes["Data Pgto Ajustada"] >= ini_mes)
        & (df_mes["Data Pgto Ajustada"] <= fim_mes)
        & (df_mes["Status Normalizado"] == "PAGO")
    ]

    total_a_pagar = df_a_pagar["Valor Original Num"].sum()
    total_baixado = df_baixado["Valor Pago Num"].sum()
    total_mes = total_a_pagar + total_baixado

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("A pagar", formatar_moeda(total_a_pagar))
    k2.metric("Pago", formatar_moeda(total_baixado))
    k3.metric("Em aberto", len(df_a_pagar))
    k4.metric("Pagos", len(df_baixado))
    st.divider()

    g1, g2 = st.columns([2, 1])

    with g1:
        st.subheader("Resumo por descrição conta")

        if df_mes.empty:
            st.info("Sem dados para os filtros selecionados.")
        else:
            resumo = (
                df_mes.groupby("Descrição conta", dropna=False)
                .agg(
                    Valor_Original=("Valor Original Num", "sum"),
                    Valor_Pago=("Valor Pago Num", "sum"),
                    Qtd=("row_id", "count"),
                )
                .reset_index()
                .sort_values("Valor_Original", ascending=False)
            )

            st.dataframe(
                resumo.head(50),
                use_container_width=True,
                height=300,
            )

    with g2:
        st.subheader("Distribuição")
        st.metric("Movimentado no filtro", formatar_moeda(total_mes))

        if total_mes > 0:
            perc_pago = (total_baixado / total_mes) * 100
        else:
            perc_pago = 0

        st.progress(min(int(perc_pago), 100), text=f"{perc_pago:.1f}% baixado")

    st.subheader("Lançamentos do filtro")

    colunas_visao = [
        COL["fornecedor"],
        COL["cnpj"],
        "Descrição conta",
        COL["vencimento"],
        COL["valor_original"],
        COL["status"],
        COL["data_pgto"],
        COL["valor_pago"],
        COL["descricao"],
    ]

    colunas_visao = [c for c in colunas_visao if c in df_mes.columns]

    st.dataframe(
        df_mes[colunas_visao].head(500),
        use_container_width=True,
        height=460,
    )

    if len(df_mes) > 500:
        st.info(f"Exibindo os primeiros 500 registros de {len(df_mes)} encontrados.")


# ============================================================
# LANÇAMENTOS E BAIXA
# ============================================================

if pagina == "Lançamentos e Baixa":
    aba1, aba2 = st.tabs(["➕ Novo lançamento", "✔ Baixa de pagamento"])

    with aba1:
        st.subheader("Novo lançamento")

        st.markdown("#### Fornecedor")

        busca_cnpj_lanc = st.text_input(
            "Pesquisar CNPJ / fornecedor",
            placeholder="Digite parte do CNPJ, razão social ou fantasia",
            key="busca_cnpj_lanc",
        )

        opcoes_cnpj = montar_opcoes_fornecedores(df_forn, busca_cnpj_lanc)

        if busca_cnpj_lanc and not opcoes_cnpj:
            st.warning("Nenhum fornecedor encontrado para a pesquisa informada.")

        if len(opcoes_cnpj) == 500:
            st.caption(
                "Exibindo os primeiros 500 fornecedores. Refine a pesquisa para localizar mais rápido."
            )

        col_cnpj, col_razao = st.columns([1, 2])

        cnpj_sel = col_cnpj.selectbox(
            "CNPJ",
            opcoes_cnpj,
            index=None,
            placeholder="Selecione o CNPJ ou fornecedor",
            key="cnpj_lancamento",
        )

        cnpj_final = ""
        fornecedor_final = ""

        if cnpj_sel:
            partes = [p.strip() for p in cnpj_sel.split("|")]
            cnpj_final = partes[1] if len(partes) > 1 else partes[0]
            fornecedor_final = obter_fornecedor_por_cnpj(df_forn, cnpj_final)

        col_razao.text_input(
            "Razão social",
            value=fornecedor_final,
            disabled=True,
        )

        col_a, col_b, col_c = st.columns(3)

        base = col_a.selectbox(
            "Base", bases_opcoes, index=None, placeholder="Selecione a base"
        )

        data_entrada = col_b.date_input("Data entrada", value=date.today())
        data_emissao = col_c.date_input("Data emissão", value=date.today())
        st.markdown("#### Plano de contas")

        busca_plano_lanc = st.text_input(
            "Pesquisar Cod_Plano Contas / descrição",
            placeholder="Digite parte do código ou da descrição",
            key="busca_plano_lanc",
        )

        opcoes_plano = montar_opcoes_plano(df_plano, busca_plano_lanc)

        if busca_plano_lanc and not opcoes_plano:
            st.warning("Nenhum plano de contas encontrado para a pesquisa informada.")

        if len(opcoes_plano) == 500:
            st.caption(
                "Exibindo os primeiros 500 planos. Refine a pesquisa para localizar mais rápido."
            )

        col_plano, col_desc = st.columns([1, 2])

        plano_sel = col_plano.selectbox(
            "Cod_Plano Contas",
            opcoes_plano,
            index=None,
            placeholder="Selecione o código ou descrição",
            key="plano_lancamento",
        )

        cod_plano_final = ""
        descricao_conta_final = ""

        if plano_sel:
            cod_plano_final = plano_sel.split("|")[0].strip()
            descricao_conta_final = obter_descricao_plano(df_plano, cod_plano_final)

        col_desc.text_input(
            "Descrição conta",
            value=descricao_conta_final,
            disabled=True,
        )

        st.markdown("#### Parcelamento")

        col_v1, col_v2 = st.columns(2)
        valor_total = col_v1.number_input("Valor total", min_value=0.0, step=0.01)
        qtd_parcelas = col_v2.number_input(
            "Quantidade de parcelas",
            min_value=1,
            max_value=60,
            value=1,
            step=1,
        )

        qtd_parcelas = int(qtd_parcelas)

        valor_base = round(valor_total / qtd_parcelas, 2) if qtd_parcelas else 0
        diferenca = round(valor_total - (valor_base * qtd_parcelas), 2)

        parcelas = []
        st.markdown("Vencimento por parcela")

        for i in range(qtd_parcelas):
            c1, c2, c3 = st.columns([1, 2, 2])
            c1.write(f"Parcela {i + 1}/{qtd_parcelas}")
            venc = c2.date_input(f"Vencimento {i + 1}", key=f"venc_{i}")

            valor_sugerido = (
                valor_base + diferenca if i == qtd_parcelas - 1 else valor_base
            )

            valor_p = c3.number_input(
                f"Valor parcela {i + 1}",
                min_value=0.0,
                value=float(valor_sugerido),
                step=0.01,
                key=f"valor_parcela_{i}",
            )

            parcelas.append(
                {
                    "numero": i + 1,
                    "vencimento": venc,
                    "valor": valor_p,
                }
            )

        soma_parcelas = round(sum(p["valor"] for p in parcelas), 2)

        if round(valor_total, 2) != soma_parcelas:
            st.warning(
                f"A soma das parcelas está diferente do valor total. "
                f"Total: {formatar_moeda(valor_total)} | Parcelas: {formatar_moeda(soma_parcelas)}"
            )

        descricao_lancamento = st.text_area("Descrição dos lançamentos", height=140)

        arquivo_origem = st.file_uploader("Anexar documento de origem", type=None)

        if st.button("Salvar lançamento", type="primary"):
            if not cnpj_final:
                st.error("Selecione um CNPJ / fornecedor.")
                st.stop()

            if not base:
                st.error("Selecione a Base.")
                st.stop()

            if not cod_plano_final:
                st.error("Selecione o Cod_Plano Contas.")
                st.stop()

            if valor_total <= 0:
                st.error("Informe um valor total maior que zero.")
                st.stop()

            if round(valor_total, 2) != soma_parcelas:
                st.error("Corrija os valores das parcelas antes de salvar.")
                st.stop()

            dados_parcelas = []

            for p in parcelas:
                dados_parcelas.append(
                    {
                        "cnpj": cnpj_final,
                        "fornecedor": fornecedor_final,
                        "base": base,
                        "entrada": str(data_entrada),
                        "emissao": str(data_emissao),
                        "vencimento": str(p["vencimento"]),
                        "cod_plano": cod_plano_final,
                        "descricao_conta": descricao_conta_final,
                        "valor_original": float(p["valor"]),
                        "status": "Aberto",
                        "descricao": descricao_lancamento,
                        "parcela": f"{qtd_parcelas:02d}/{p['numero']:02d}",
                    }
                )

            with st.spinner("Gravando no Smartsheet..."):
                retorno = criar_lancamentos(col_map_contas, dados_parcelas)

                if arquivo_origem:
                    linhas_criadas = retorno.get("result", [])
                    for linha in linhas_criadas:
                        upload_anexo(linha["id"], arquivo_origem)

            st.cache_data.clear()
            st.success("Lançamento salvo com sucesso.")
            st.rerun()

    with aba2:
        st.subheader("Baixa de pagamento")

        if df_contas.empty:
            st.warning("Nenhum lançamento encontrado.")
            st.stop()

        df_baixa = df_contas[df_contas["Status Normalizado"] != "PAGO"].copy()

        colf1, colf2, colf3 = st.columns(3)
        busca = colf1.text_input("Pesquisar lançamento para baixa")
        ano_baixa = colf2.selectbox(
            "Ano vencimento",
            ["Todos"] + list(range(date.today().year - 5, date.today().year + 3)),
        )
        mes_baixa = colf3.selectbox("Mês vencimento", ["Todos"] + list(range(1, 13)))

        if ano_baixa != "Todos":
            df_baixa = df_baixa[
                pd.to_datetime(
                    df_baixa["Data Vencimento Ajustada"], errors="coerce"
                ).dt.year
                == int(ano_baixa)
            ]

        if mes_baixa != "Todos":
            df_baixa = df_baixa[
                pd.to_datetime(
                    df_baixa["Data Vencimento Ajustada"], errors="coerce"
                ).dt.month
                == int(mes_baixa)
            ]

        if busca:
            mask = texto_pesquisa_df(df_baixa).str.contains(
                busca.upper(), na=False, regex=False
            )
            df_baixa = df_baixa[mask]

        df_baixa = df_baixa.sort_values(
            "Data Vencimento Ajustada", ascending=True
        ).head(300)

        if df_baixa.empty:
            st.info("Nenhum lançamento em aberto encontrado para os filtros.")
            st.stop()

        label = (
            df_baixa["row_id"].astype(str)
            + " | "
            + df_baixa.get(COL["fornecedor"], "").fillna("").astype(str)
            + " | Venc: "
            + df_baixa.get(COL["vencimento"], "").fillna("").astype(str)
            + " | Valor: "
            + df_baixa.get(COL["valor_original"], "").fillna("").astype(str)
        )

        opcoes_baixa = label.tolist()

        selecionado = st.selectbox(
            "Lançamento",
            opcoes_baixa,
            index=None,
            placeholder="Pesquise/Selecione o lançamento para baixar",
        )

        if selecionado:
            row_id = selecionado.split("|")[0].strip()
            linha = df_baixa[df_baixa["row_id"].astype(str) == row_id]

            valor_sugerido = 0.0
            if not linha.empty:
                valor_sugerido = float(linha.iloc[0].get("Valor Original Num", 0.0))

            colp1, colp2 = st.columns(2)
            valor_pago = colp1.number_input(
                "Valor pago",
                min_value=0.0,
                value=valor_sugerido,
                step=0.01,
            )
            data_pgto = colp2.date_input("Data Pgto", value=date.today())

            comprovante = st.file_uploader("Anexar comprovante de pagamento (opcional)")

            if st.button("Confirmar baixa", type="primary"):
                with st.spinner("Registrando baixa no Smartsheet..."):
                    marcar_pago(row_id, col_map_contas, valor_pago, data_pgto)

                    if comprovante:
                        upload_anexo(row_id, comprovante)

                st.cache_data.clear()
                st.success("Baixa registrada com sucesso.")
                st.rerun()
