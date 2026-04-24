import streamlit as st
import requests
import pandas as pd
from datetime import date
from calendar import monthrange

SMARTSHEET_TOKEN = "Wmw5xFBAyj82NOIYU2yPObLR8FtRXm6c6QgV4"

SHEETS = {
    "contas": "6432342517540740",
    "fornecedores": "4516477865879428",
    "plano_contas": "8420259540559748",
}

BASE_URL = "https://api.smartsheet.com/2.0"

HEADERS = {
    "Authorization": f"Bearer {SMARTSHEET_TOKEN}",
    "Content-Type": "application/json",
}

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


def api_get(path):
    r = requests.get(f"{BASE_URL}{path}", headers=HEADERS)
    if r.status_code != 200:
        st.error(f"Erro na API: {r.text}")
        st.stop()
    return r.json()


def api_post(path, payload):
    r = requests.post(f"{BASE_URL}{path}", headers=HEADERS, json=payload)
    if r.status_code not in [200, 201]:
        st.error(f"Erro ao gravar: {r.text}")
        st.stop()
    return r.json()


def api_put(path, payload):
    r = requests.put(f"{BASE_URL}{path}", headers=HEADERS, json=payload)
    if r.status_code not in [200, 201]:
        st.error(f"Erro ao atualizar: {r.text}")
        st.stop()
    return r.json()


@st.cache_data(ttl=300, show_spinner=False)
def carregar_sheet(sheet_id):
    return api_get(f"/sheets/{sheet_id}")


def mapa_colunas(sheet):
    return {c["title"]: c["id"] for c in sheet["columns"]}


def sheet_para_dataframe(sheet):
    col_id_para_nome = {c["id"]: c["title"] for c in sheet["columns"]}
    linhas = []

    for row in sheet.get("rows", []):
        item = {"row_id": row["id"]}
        for cell in row.get("cells", []):
            titulo = col_id_para_nome.get(cell["columnId"])
            if titulo:
                item[titulo] = cell.get("displayValue", cell.get("value"))
        linhas.append(item)

    return pd.DataFrame(linhas), mapa_colunas(sheet)


def mapa_colunas(sheet):
    return {c["title"]: c["id"] for c in sheet["columns"]}


def sheet_para_dataframe(sheet):
    colunas = mapa_colunas(sheet)
    linhas = []

    for row in sheet.get("rows", []):
        item = {"row_id": row["id"]}
        for cell in row.get("cells", []):
            titulo = next(
                (k for k, v in colunas.items() if v == cell["columnId"]), None
            )
            if titulo:
                item[titulo] = cell.get("displayValue", cell.get("value"))
        linhas.append(item)

    return pd.DataFrame(linhas), colunas


def valor_numero(v):
    if pd.isna(v) or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    v = str(v).replace("R$", "").replace(".", "").replace(",", ".").strip()
    try:
        return float(v)
    except:
        return 0.0


def data_segura(v):
    try:
        return pd.to_datetime(v).date()
    except:
        return None


def criar_lancamentos(col_map, dados_parcelas):
    rows = []

    for item in dados_parcelas:
        cells = []

        for campo, valor in item.items():
            nome_coluna = COL.get(campo)
            if nome_coluna and nome_coluna in col_map:
                cells.append({"columnId": col_map[nome_coluna], "value": valor})

        rows.append({"toBottom": True, "cells": cells})

    return api_post(f"/sheets/{SHEETS['contas']}/rows", rows)


def marcar_pago(row_id, col_map, valor_pago, data_pgto):
    cells = [
        {"columnId": col_map[COL["status"]], "value": "Pago"},
        {"columnId": col_map[COL["valor_pago"]], "value": float(valor_pago)},
        {"columnId": col_map[COL["data_pgto"]], "value": str(data_pgto)},
    ]

    payload = {"rows": [{"id": int(row_id), "cells": cells}]}

    return api_put(f"/sheets/{SHEETS['contas']}/rows", payload)


def upload_anexo(row_id, arquivo):
    url = f"{BASE_URL}/sheets/{SHEETS['contas']}/rows/{row_id}/attachments"
    headers = {"Authorization": f"Bearer {SMARTSHEET_TOKEN}"}
    files = {"file": (arquivo.name, arquivo.getvalue())}

    r = requests.post(url, headers=headers, files=files)
    if r.status_code not in [200, 201]:
        st.error(f"Erro ao anexar arquivo: {r.text}")
        st.stop()


def formatar_moeda(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


st.set_page_config(page_title="Contas a Pagar", layout="wide")

st.title("💰 Contas a Pagar")

sheet_contas = carregar_sheet(SHEETS["contas"])
df_contas, col_map_contas = sheet_para_dataframe(sheet_contas)

sheet_fornecedores = carregar_sheet(SHEETS["fornecedores"])
df_forn, _ = sheet_para_dataframe(sheet_fornecedores)

sheet_plano = carregar_sheet(SHEETS["plano_contas"])
df_plano, _ = sheet_para_dataframe(sheet_plano)

pagina = st.sidebar.radio("Menu", ["Dashboard", "Lançamentos e Baixa"])

if not df_contas.empty:
    df_contas["Valor Original Num"] = df_contas.get(COL["valor_original"], 0).apply(
        valor_numero
    )
    df_contas["Valor Pago Num"] = df_contas.get(COL["valor_pago"], 0).apply(
        valor_numero
    )
    df_contas["Data Vencimento Ajustada"] = df_contas.get(
        COL["vencimento"], None
    ).apply(data_segura)
    df_contas["Data Pgto Ajustada"] = df_contas.get(COL["data_pgto"], None).apply(
        data_segura
    )
    df_contas["Descrição conta"] = df_contas.get(COL["descricao_conta"], "")

if pagina == "Dashboard":
    hoje = date.today()

    col1, col2, col3, col4 = st.columns(4)

    ano = col1.selectbox("Ano", list(range(hoje.year - 3, hoje.year + 2)), index=3)

    mes = col2.selectbox("Mês", list(range(1, 13)), index=hoje.month - 1)

    if df_contas.empty:
        st.warning("Nenhum lançamento encontrado.")
    else:
        df_dash = df_contas.copy()

        df_dash["Valor Original Num"] = df_dash.get(COL["valor_original"], 0).apply(
            valor_numero
        )
        df_dash["Valor Pago Num"] = df_dash.get(COL["valor_pago"], 0).apply(
            valor_numero
        )
        df_dash["Data Vencimento Ajustada"] = df_dash.get(
            COL["vencimento"], None
        ).apply(data_segura)
        df_dash["Data Pgto Ajustada"] = df_dash.get(COL["data_pgto"], None).apply(
            data_segura
        )
        df_dash["Descrição conta"] = df_dash.get(COL["descricao_conta"], "")

        ini_mes = date(ano, mes, 1)
        fim_mes = date(ano, mes, monthrange(ano, mes)[1])

        # Primeiro limita ao mês escolhido para melhorar performance
        df_mes = df_dash[
            (
                (df_dash["Data Vencimento Ajustada"].notna())
                & (df_dash["Data Vencimento Ajustada"] >= ini_mes)
                & (df_dash["Data Vencimento Ajustada"] <= fim_mes)
            )
            | (
                (df_dash["Data Pgto Ajustada"].notna())
                & (df_dash["Data Pgto Ajustada"] >= ini_mes)
                & (df_dash["Data Pgto Ajustada"] <= fim_mes)
            )
        ].copy()

        descricoes = ["Todas"]
        if "Descrição conta" in df_mes.columns:
            descricoes += sorted(
                df_mes["Descrição conta"].dropna().astype(str).unique().tolist()
            )

        filtro_descricao = col3.selectbox("Descrição conta", descricoes)

        status_opcoes = ["Todos"]
        if COL["status"] in df_mes.columns:
            status_opcoes += sorted(
                df_mes[COL["status"]].dropna().astype(str).unique().tolist()
            )

        filtro_status = col4.selectbox("Status", status_opcoes)

        if filtro_descricao != "Todas":
            df_mes = df_mes[df_mes["Descrição conta"].astype(str) == filtro_descricao]

        if filtro_status != "Todos":
            df_mes = df_mes[df_mes[COL["status"]].astype(str) == filtro_status]

        df_a_pagar = df_mes[
            (df_mes["Data Vencimento Ajustada"].notna())
            & (df_mes["Data Vencimento Ajustada"] >= ini_mes)
            & (df_mes["Data Vencimento Ajustada"] <= fim_mes)
            & (df_mes.get(COL["status"], "").astype(str).str.upper() != "PAGO")
        ]

        df_baixado = df_mes[
            (df_mes["Data Pgto Ajustada"].notna())
            & (df_mes["Data Pgto Ajustada"] >= ini_mes)
            & (df_mes["Data Pgto Ajustada"] <= fim_mes)
            & (df_mes.get(COL["status"], "").astype(str).str.upper() == "PAGO")
        ]

        total_a_pagar = df_a_pagar["Valor Original Num"].sum()
        total_baixado = df_baixado["Valor Pago Num"].sum()

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("A pagar no mês", formatar_moeda(total_a_pagar))
        k2.metric("Baixado no mês", formatar_moeda(total_baixado))
        k3.metric("Qtd. em aberto", len(df_a_pagar))
        k4.metric("Qtd. baixados", len(df_baixado))

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
            df_mes[colunas_visao].head(500), use_container_width=True, height=480
        )

        if len(df_mes) > 500:
            st.info(
                f"Exibindo os primeiros 500 registros de {len(df_mes)} encontrados."
            )

if pagina == "Lançamentos e Baixa":
    aba1, aba2 = st.tabs(["Novo lançamento", "Baixa de pagamento"])

    with aba1:
        st.subheader("Novo lançamento")

        cnpj_col = COL["cnpj"] if COL["cnpj"] in df_forn.columns else None
        forn_col = COL["fornecedor"] if COL["fornecedor"] in df_forn.columns else None
        fantasia_col = COL["fantasia"] if COL["fantasia"] in df_forn.columns else None

        if cnpj_col:
            busca_cnpj = st.text_input("Pesquisar CNPJ / fornecedor")
            df_filtro_forn = df_forn.copy()

            if busca_cnpj:
                texto = busca_cnpj.upper()
                df_filtro_forn = df_filtro_forn[
                    df_filtro_forn.astype(str)
                    .apply(lambda x: x.str.upper().str.contains(texto, na=False))
                    .any(axis=1)
                ]

            opcoes_cnpj = []
        for _, r in df_filtro_forn.iterrows():
            cnpj = str(r.get(cnpj_col, "")).strip()
            fornecedor = str(r.get(forn_col, "")).strip() if forn_col else ""
            fantasia = str(r.get(fantasia_col, "")).strip() if fantasia_col else ""

            label = f"{cnpj} | {fornecedor} | {fantasia}"
            opcoes_cnpj.append(label)

        cnpj_sel = (
            st.selectbox(
                "CNPJ / Fornecedor",
                opcoes_cnpj,
                index=None,
                placeholder="Pesquise por CNPJ, fornecedor ou fantasia",
            )
            if opcoes_cnpj
            else None
        )

        col_a, col_b, col_c = st.columns(3)
        base = col_a.text_input("Base")
        data_entrada = col_b.date_input("Data entrada")
        data_emissao = col_c.date_input("Data emissão")

        st.markdown("#### Plano de contas")

        opcoes_plano = montar_opcoes_plano(df_plano, "")

        col_plano, col_desc = st.columns([1, 2])

        plano_sel = col_plano.selectbox(
            "Cod_Plano Contas",
            opcoes_plano,
            index=None,
            placeholder="Pesquise e selecione o código ou descrição",
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
            "Quantidade de parcelas", min_value=1, max_value=60, value=1, step=1
        )

        valor_parcela = round(valor_total / qtd_parcelas, 2) if qtd_parcelas else 0

        parcelas = []
        st.markdown("Vencimento por parcela")

        for i in range(int(qtd_parcelas)):
            c1, c2, c3 = st.columns([1, 2, 2])
            c1.write(f"Parcela {i + 1}/{int(qtd_parcelas)}")
            venc = c2.date_input(f"Vencimento {i + 1}", key=f"venc_{i}")
            valor_p = c3.number_input(
                f"Valor parcela {i + 1}",
                min_value=0.0,
                value=valor_parcela,
                step=0.01,
                key=f"valor_parcela_{i}",
            )

            parcelas.append({"numero": i + 1, "vencimento": venc, "valor": valor_p})

        descricao_lancamento = st.text_area("Descrição dos lançamentos", height=140)

        arquivo_origem = st.file_uploader("Anexar documento de origem", type=None)

        if st.button("Salvar lançamento"):
            if not cnpj_final:
                st.error("Selecione um CNPJ.")
                st.stop()

            if not cod_plano_final:
                st.error("Selecione o Cod_Plano Contas.")
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
                        "parcela": f"{p['numero']}/{int(qtd_parcelas)}",
                    }
                )

            retorno = criar_lancamentos(col_map_contas, dados_parcelas)

            if arquivo_origem:
                linhas_criadas = retorno.get("result", [])
                for linha in linhas_criadas:
                    upload_anexo(linha["id"], arquivo_origem)

            st.cache_data.clear()
            st.success("Lançamento salvo com sucesso.")

    with aba2:
        st.subheader("Baixa de pagamento")

        if df_contas.empty:
            st.warning("Nenhum lançamento encontrado.")
        else:
            df_baixa = df_contas.copy()

            if COL["status"] in df_baixa.columns:
                df_baixa = df_baixa[
                    df_baixa[COL["status"]].astype(str).str.upper() != "PAGO"
                ]

            busca = st.text_input("Pesquisar lançamento para baixa")

            if busca:
                texto = busca.upper()
                df_baixa = df_baixa[
                    df_baixa.astype(str)
                    .apply(lambda x: x.str.upper().str.contains(texto, na=False))
                    .any(axis=1)
                ]

            opcoes = []
            for _, r in df_baixa.iterrows():
                fornecedor = r.get(COL["fornecedor"], "")
                venc = r.get(COL["vencimento"], "")
                valor = r.get(COL["valor_original"], "")
                row_id = r.get("row_id")
                opcoes.append(
                    f"{row_id} | {fornecedor} | Venc: {venc} | Valor: {valor}"
                )

            selecionado = st.selectbox("Lançamento", opcoes) if opcoes else None

            if selecionado:
                row_id = selecionado.split("|")[0].strip()

                colp1, colp2 = st.columns(2)
                valor_pago = colp1.number_input("Valor pago", min_value=0.0, step=0.01)
                data_pgto = colp2.date_input("Data pagamento", value=date.today())

                comprovante = st.file_uploader("Anexar comprovante de pagamento")

                if st.button("Confirmar baixa"):
                    marcar_pago(row_id, col_map_contas, valor_pago, data_pgto)

                    if comprovante:
                        upload_anexo(row_id, comprovante)

                    st.cache_data.clear()
                    st.success("Baixa registrada com sucesso.")
