import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
import random
import re
import matplotlib.pyplot as plt

# Biblioteca para conexão com SharePoint
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.user_credential import UserCredential

###############################################################################
# 1) CREDENCIAIS E CAMINHOS NO SHAREPOINT
###############################################################################
EMAIL_REMETENTE = st.secrets["sharepoint"]["email"]
SENHA_EMAIL = st.secrets["sharepoint"]["password"]
SITE_URL = st.secrets["sharepoint"]["site_url"]

FILE_URL_FORNECEDORES = st.secrets["sharepoint"]["file_url"]  # Excel de Fornecedores
FILE_URL_MENSAL_2025 = "/sites/gestaodeprodutos/Documentos Compartilhados/Gestão financeira/Controle mensal de pagamento - 2025 (novo) - Automation.xlsx"
FILE_URL_MENSAL_2026 = "/sites/gestaodeprodutos/Documentos Compartilhados/Gestão financeira/Controle mensal de pagamento - 2026 (novo) - Automation.xlsx"

###############################################################################
# 2) DEFINIÇÃO DE COLUNAS E LISTAS
###############################################################################
GENERAL_COLUMNS = [
    "Fornecedor",
    "ID - Fornecedor",
    "CNPJ",
    "Contato",
    "Centro de custo",
]

SPECIFIC_COLUMNS = [
    "Nº do Serviço",
    "ID - Produto",
    "Categoria do Produto",
    "Descrição do Produto",
    "Localidade",
    "Status",
    "Inicio do contrato",
    "Termino do contrato",
    "Tempo do contrato",
    "Metodo de pagamento",
    "Tipo de pagamento",
    "Dia de Pagamento",
    "ID - Pagamento",
    "Status de Pagamento",
    "Valor mensal",
    "Valor do plano",
    "Tempo de pagamento",
    "Orçado",
    "Observações",
    "Forma de pagamento",
    "Início do Pagamento",
]

ALL_COLUMNS = GENERAL_COLUMNS + SPECIFIC_COLUMNS

category_options = [
    "Internet Link",
    "Segurança",
    "Telefonia",
    "Impressão",
    "Licença Office",
    "Software",
    "Suporte",
    "Serviço - Atendimento",
    "Software - Atendimento",
    "Hardware",
    "Cloud",
    "Material de infra",
]

COLUNAS_CONTROLE_MENSAL = [
    "Fornecedor",
    "ID - Fornecedor",
    "ID - Pagamento",
    "Categoria",
    "Dia Vencimento",
    "Data Envio",
    "Data Pagamento",
    "Metodo de Pagamento",
    "Status de Pagamento",
    "Planejado",
    "Moeda",
    "Valor Estimado - Real",
    "Valor Pago Convertido",
    "Diferença",
    "Observações",
    "Ano",
    "Mes",
]

MESES_ORDENADOS = [
    "JANEIRO",
    "FEVEREIRO",
    "MARCO",
    "ABRIL",
    "MAIO",
    "JUNHO",
    "JULHO",
    "AGOSTO",
    "SETEMBRO",
    "OUTUBRO",
    "NOVEMBRO",
    "DEZEMBRO",
]

MOEDAS_COMUNS = ["REAL", "DOLAR", "EURO"]
STATUS_PAG_OPCOES = ["PENDENTE", "PAGO"]

###############################################################################
# 3) FUNÇÕES AUXILIARES
###############################################################################
def parse_float_br(valor_str):
    """
    Converte string estilo 'R$ 1.234,56' -> 1234.56 (float).
    """
    if not isinstance(valor_str, str):
        return valor_str
    v = valor_str.replace("R$", "").strip()
    if "." in v and "," in v:
        v = v.replace(".", "")
        v = v.replace(",", ".")
    elif "," in v and "." not in v:
        v = v.replace(",", ".")
    try:
        return float(v)
    except ValueError:
        return None

def _datetime_to_str(data):
    """
    Converte datetime/date -> 'DD/MM/AAAA'.
    """
    if pd.isnull(data):
        return ""
    if isinstance(data, (pd.Timestamp, datetime.date, datetime.datetime)):
        return data.strftime("%d/%m/%Y")
    return str(data)

def parse_date_br(data_str):
    """
    Tenta converter 'DD/MM/AAAA' -> datetime.date (ou None).
    """
    if not data_str.strip():
        return None
    try:
        return datetime.datetime.strptime(data_str.strip(), "%d/%m/%Y").date()
    except ValueError:
        return None

def generate_id_fornecedor(nome):
    """
    Gera ID do fornecedor ex: 'SYN123'.
    """
    if not nome:
        return ""
    prefix = re.sub(r"[^A-Za-z]", "", nome).upper()[:3]
    rand_num = random.randint(100, 999)
    return f"{prefix}{rand_num}"

def generate_id_produto(descricao, categoria):
    """
    Gera ID do produto, ex: 'INTTEL456'.
    """
    if not descricao or not categoria:
        return ""
    d = re.sub(r"[^A-Za-z]", "", descricao).upper()[:3]
    c = re.sub(r"[^A-Za-z]", "", categoria).upper()[:3]
    rand_num = random.randint(100, 999)
    return f"{d}{c}{rand_num}"

###############################################################################
# 4) CARREGAR EXCEL DO SHAREPOINT
###############################################################################
from office365.sharepoint.files.file import File

def load_excel_from_sharepoint(file_url):
    ctx = ClientContext(SITE_URL).with_credentials(UserCredential(EMAIL_REMETENTE, SENHA_EMAIL))
    response = File.open_binary(ctx, file_url)
    excel_data = response.content
    sheets = pd.read_excel(BytesIO(excel_data), sheet_name=None)
    return sheets

###############################################################################
# 5) CÓDIGO PARA FORNECEDORES
###############################################################################
def load_fornecedores():
    try:
        all_sheets = load_excel_from_sharepoint(FILE_URL_FORNECEDORES)
        results = {}
        for sheet_name, df in all_sheets.items():
            df.columns = df.columns.str.strip()
            for col in ALL_COLUMNS:
                if col not in df.columns:
                    df[col] = ""
            if "CNPJ" in df.columns:
                df["CNPJ"] = df["CNPJ"].astype(str)
            if "Contato" in df.columns:
                df["Contato"] = df["Contato"].astype(str)
            if "Inicio do contrato" in df.columns:
                df["Inicio do contrato"] = df["Inicio do contrato"].apply(_datetime_to_str)
            if "Termino do contrato" in df.columns:
                df["Termino do contrato"] = df["Termino do contrato"].apply(_datetime_to_str)
            if "Início do Pagamento" in df.columns:
                df["Início do Pagamento"] = df["Início do Pagamento"].apply(_datetime_to_str)
            if "Valor mensal" in df.columns:
                df["Valor mensal"] = df["Valor mensal"].apply(parse_float_br)
            if "Valor do plano" in df.columns:
                df["Valor do plano"] = df["Valor do plano"].apply(parse_float_br)
            results[sheet_name] = df
        return results
    except Exception as e:
        st.error(f"Erro ao carregar Fornecedores: {e}")
        return {}

def save_fornecedores():
    try:
        for sheet_name, df in st.session_state.suppliers_data.items():
            if "CNPJ" in df.columns:
                df["CNPJ"] = df["CNPJ"].astype(str)
            if "Contato" in df.columns:
                df["Contato"] = df["Contato"].astype(str)
            if "Inicio do contrato" in df.columns:
                df["Inicio do contrato"] = df["Inicio do contrato"].apply(_datetime_to_str)
            if "Termino do contrato" in df.columns:
                df["Termino do contrato"] = df["Termino do contrato"].apply(_datetime_to_str)
            if "Início do Pagamento" in df.columns:
                df["Início do Pagamento"] = df["Início do Pagamento"].apply(_datetime_to_str)

        output = BytesIO()
        with pd.ExcelWriter(output) as writer:
            for sheet_name, df in st.session_state.suppliers_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        ctx = ClientContext(SITE_URL).with_credentials(UserCredential(EMAIL_REMETENTE, SENHA_EMAIL))
        File.save_binary(ctx, FILE_URL_FORNECEDORES, output.getvalue())
        st.success("Dados de Fornecedores salvos com sucesso! Para visualizar, atualize a página ou acesse a aba 'Lista de Fornecedores'.")
    except Exception as e:
        if "Locked" in str(e) or "423" in str(e):
            st.warning("Arquivo de Fornecedores bloqueado. Feche ou faça check-in antes de salvar.")
        else:
            st.error(f"Erro ao salvar Fornecedores: {e}")

###############################################################################
# 6) CÓDIGO PARA CONTROLE MENSAL
###############################################################################
def load_controle_mensal():
    """
    Lê os arquivos (2025, 2026), ignora 'MATRIZ' e abas fora de MESES_ORDENADOS.
    Concatena num único DataFrame.
    """
    dfs = []
    map_ano_arquivo = {
        "2025": FILE_URL_MENSAL_2025,
        "2026": FILE_URL_MENSAL_2026,
    }
    for ano, url_arq in map_ano_arquivo.items():
        try:
            sheets = load_excel_from_sharepoint(url_arq)
            for sheet_name, df_mes in sheets.items():
                sn = sheet_name.strip().upper()
                if sn == "MATRIZ":
                    continue
                if sn not in MESES_ORDENADOS:
                    continue
                df_mes.columns = df_mes.columns.str.strip()
                for col in COLUNAS_CONTROLE_MENSAL:
                    if col not in df_mes.columns:
                        df_mes[col] = ""
                if "Data Envio" in df_mes.columns:
                    df_mes["Data Envio"] = pd.to_datetime(df_mes["Data Envio"], errors="coerce", dayfirst=True)
                if "Data Pagamento" in df_mes.columns:
                    df_mes["Data Pagamento"] = pd.to_datetime(df_mes["Data Pagamento"], errors="coerce", dayfirst=True)
                if "Valor Estimado - Real" in df_mes.columns:
                    df_mes["Valor Estimado - Real"] = df_mes["Valor Estimado - Real"].astype(str).apply(parse_float_br)
                if "Valor Pago Convertido" in df_mes.columns:
                    df_mes["Valor Pago Convertido"] = df_mes["Valor Pago Convertido"].astype(str).apply(parse_float_br)

                df_mes["Ano"] = ano
                df_mes["Mes"] = sn
                dfs.append(df_mes)
        except Exception as e:
            st.warning(f"Erro ao carregar {url_arq} ({ano}): {e}")

    if len(dfs) == 0:
        return pd.DataFrame(columns=COLUNAS_CONTROLE_MENSAL)
    df_final = pd.concat(dfs, ignore_index=True)
    df_final["Mes_Indice"] = df_final["Mes"].apply(lambda x: MESES_ORDENADOS.index(x) if x in MESES_ORDENADOS else 99)
    df_final = df_final.sort_values(by=["Ano", "Mes_Indice"]).drop(columns=["Mes_Indice"])
    return df_final

def save_controle_mensal():
    """
    Salva st.session_state["controle_mensal"] particionado por (Ano, Mes).
    Apenas mostra mensagem de sucesso para cada ano que for realmente salvo.
    """
    df = st.session_state["controle_mensal"].copy()
    if df.empty:
        st.warning("Não há pagamentos para salvar.")
        return

    group = df.groupby(["Ano", "Mes"], as_index=False)
    dict_ano_mes = {}
    for (ano, mes), df_subset in group:
        for col in COLUNAS_CONTROLE_MENSAL:
            if col not in df_subset.columns:
                df_subset[col] = ""
        if "Data Envio" in df_subset.columns:
            df_subset["Data Envio"] = df_subset["Data Envio"].apply(_datetime_to_str)
        if "Data Pagamento" in df_subset.columns:
            df_subset["Data Pagamento"] = df_subset["Data Pagamento"].apply(_datetime_to_str)
        if ano not in dict_ano_mes:
            dict_ano_mes[ano] = {}
        dict_ano_mes[ano][mes] = df_subset

    map_ano_arquivo = {
        "2025": FILE_URL_MENSAL_2025,
        "2026": FILE_URL_MENSAL_2026,
    }

    for ano, meses_dict in dict_ano_mes.items():
        if ano not in map_ano_arquivo:
            st.warning(f"Ano {ano} não mapeado. Ignorando.")
            continue

        output = BytesIO()
        with pd.ExcelWriter(output) as writer:
            # Garante a ordem de Janeiro a Dezembro nas abas
            for mes in MESES_ORDENADOS:
                if mes in meses_dict:
                    df_abames = meses_dict[mes]
                    df_abames.to_excel(writer, sheet_name=mes, index=False)

        try:
            ctx = ClientContext(SITE_URL).with_credentials(UserCredential(EMAIL_REMETENTE, SENHA_EMAIL))
            File.save_binary(ctx, map_ano_arquivo[ano], output.getvalue())
            # Mensagem de sucesso mais clara:
            st.success(f"Os pagamentos referentes a {ano} foram salvos com sucesso no Excel do SharePoint!")
        except Exception as e:
            if "Locked" in str(e) or "423" in str(e):
                st.warning("Arquivo de Pagamentos bloqueado. Feche ou faça check-in antes de salvar.")
            else:
                st.error(f"Erro ao salvar pagamentos de {ano}: {e}")

###############################################################################
# 7) INICIALIZA ST.SESSION_STATE
###############################################################################
if "suppliers_data" not in st.session_state:
    st.session_state.suppliers_data = load_fornecedores()

if "controle_mensal" not in st.session_state:
    st.session_state["controle_mensal"] = load_controle_mensal()

if "fornecedor_criado" not in st.session_state:
    st.session_state.fornecedor_criado = False

###############################################################################
# 8) CRIA AS ABAS NO STREAMLIT
###############################################################################
tabs = st.tabs([
    "Gerenciar Fornecedores",
    "Lista de Fornecedores",
    "Registrar Pagamentos"  # Nome da aba ajustado
])

###############################################################################
# ABA 1: GERENCIAR FORNECEDORES
###############################################################################
with tabs[0]:
    st.title("Gerenciar Fornecedores")

    suppliers = list(st.session_state.suppliers_data.keys())
    selected_supplier = st.sidebar.selectbox(
        "Selecione o Fornecedor",
        ["Adicionar Novo Fornecedor"] + suppliers
    )

    def _auto_calc_valor_plano():
        val_mensal = parse_float_br(st.session_state.get("new_valor_mensal_str", ""))
        try:
            parc = int(st.session_state.get("new_tempo_pagamento", "1"))
        except ValueError:
            parc = 1
        if val_mensal is not None and parc > 0:
            st.session_state["new_valor_plano_str"] = f"{val_mensal * parc:.2f}"
        else:
            st.session_state["new_valor_plano_str"] = ""

    # Inicia variáveis se não existem
    for var_key in ["new_valor_mensal_str", "new_tempo_pagamento", "new_valor_plano_str"]:
        if var_key not in st.session_state:
            st.session_state[var_key] = ""

    if "supplier_name" not in st.session_state:
        st.session_state.supplier_name = ""
    if "product_name" not in st.session_state:
        st.session_state.product_name = ""

    if st.session_state.fornecedor_criado and selected_supplier == "Adicionar Novo Fornecedor":
        st.success("Fornecedor criado com sucesso! Atualize a página para visualizar ou acesse a aba 'Lista de Fornecedores'.")
        if st.button("Criar outro fornecedor", key="criar_outro_fornecedor"):
            st.session_state.fornecedor_criado = False
            st.session_state["new_valor_mensal_str"] = ""
            st.session_state["new_tempo_pagamento"] = ""
            st.session_state["new_valor_plano_str"] = ""
            st.session_state["supplier_name"] = ""
            st.session_state["product_name"] = ""
            st.session_state["auto_id_fornecedor"] = ""
            st.session_state["auto_id_produto"] = ""
        else:
            st.info("Clique em 'Criar outro fornecedor' para limpar o formulário.")

    elif selected_supplier == "Adicionar Novo Fornecedor" and not st.session_state.fornecedor_criado:
        st.subheader("Adicionar Novo Fornecedor")

        new_supplier_name = st.text_input("Nome do Fornecedor", value=st.session_state.supplier_name, key="novo_fornecedor_nome")
        if new_supplier_name != st.session_state.supplier_name:
            st.session_state.supplier_name = new_supplier_name
            st.session_state.auto_id_fornecedor = generate_id_fornecedor(new_supplier_name)

        new_id_fornecedor = st.text_input("ID - Fornecedor (auto)",
                                          value=st.session_state.get("auto_id_fornecedor", ""),
                                          key="novo_fornecedor_id")

        new_cnpj = st.text_input("CNPJ", key="novo_fornecedor_cnpj")
        new_contato = st.text_input("Contato", key="novo_fornecedor_contato")
        new_centro_custo = st.text_input("Centro de custo", key="novo_fornecedor_custo")

        new_descricao_produto = st.text_input("Descrição do Produto",
                                              value=st.session_state.product_name,
                                              key="novo_fornecedor_desc")
        if new_descricao_produto != st.session_state.product_name:
            st.session_state.product_name = new_descricao_produto

        new_categoria_produto = st.selectbox("Categoria do Produto", category_options, key="novo_fornecedor_categoria")
        st.session_state.auto_id_produto = generate_id_produto(new_descricao_produto, new_categoria_produto)
        new_id_produto = st.text_input("ID - Produto (auto)",
                                       value=st.session_state.get("auto_id_produto", ""),
                                       key="novo_produto_id")

        new_status = "ATIVO"
        localidades = ["PAULINIA", "AMBAS", "CAMPINAS"]
        new_localidade = st.selectbox("Localidade", localidades, key="novo_fornecedor_localidade")
        new_metodo_pagamento = st.selectbox("Método de Pagamento", ["BOLETO", "CARTÃO"], key="novo_fornecedor_metodopag")

        forma_options = ["A Prazo", "A Vista"]
        new_forma_pagamento = st.selectbox("Forma de pagamento", forma_options, key="novo_fornecedor_formapag")
        if new_forma_pagamento == "A Vista":
            st.session_state["new_tempo_pagamento"] = "1"

        st.text_input("Valor Mensal (R$)", key="new_valor_mensal_str", on_change=_auto_calc_valor_plano)
        st.text_input("Tempo de pagamento (Parcelas)", key="new_tempo_pagamento", on_change=_auto_calc_valor_plano)
        st.text_input("Valor do Plano (R$) - Autopreenchido", key="new_valor_plano_str", disabled=True)

        new_inicio_str = st.text_input("Início do contrato (DD/MM/AAAA)", key="novo_fornecedor_inicio")
        new_termino_str = st.text_input("Término do contrato (DD/MM/AAAA)", key="novo_fornecedor_termino")
        new_inicio_pag_str = st.text_input("Início do Pagamento (DD/MM/AAAA)", key="novo_fornecedor_iniciopag")

        new_orcado = st.selectbox("Orçado", ["Sim", "Não"], key="novo_fornecedor_orcado")
        new_observacoes = st.text_input("Observações", "", key="novo_fornecedor_observacoes")

        if st.button("Criar Fornecedor", key="botao_criar_fornecedor"):
            val_mensal = parse_float_br(st.session_state["new_valor_mensal_str"])
            val_plano = parse_float_br(st.session_state["new_valor_plano_str"])

            dt_inicio = parse_date_br(new_inicio_str) or None
            dt_termino = parse_date_br(new_termino_str) or None
            dt_inicio_pag = parse_date_br(new_inicio_pag_str) or None

            tempo_contrato = 0
            if dt_inicio and dt_termino:
                meses = (dt_termino.year - dt_inicio.year) * 12 + (dt_termino.month - dt_inicio.month)
                if dt_termino.day >= dt_inicio.day:
                    meses += 1
                if meses < 0:
                    meses = 0
                tempo_contrato = meses

            new_row = {
                "Fornecedor": new_supplier_name,
                "ID - Fornecedor": new_id_fornecedor,
                "CNPJ": new_cnpj,
                "Contato": new_contato,
                "Centro de custo": new_centro_custo,
                "Descrição do Produto": new_descricao_produto,
                "ID - Produto": new_id_produto,
                "Categoria do Produto": new_categoria_produto,
                "Status": new_status,
                "Localidade": new_localidade,
                "Metodo de pagamento": new_metodo_pagamento,
                "Forma de pagamento": new_forma_pagamento,
                "Valor mensal": val_mensal,
                "Valor do plano": val_plano,
                "Tempo de pagamento": st.session_state["new_tempo_pagamento"],
                "Inicio do contrato": _datetime_to_str(dt_inicio),
                "Termino do contrato": _datetime_to_str(dt_termino),
                "Tempo do contrato": tempo_contrato,
                "Início do Pagamento": _datetime_to_str(dt_inicio_pag),
                "Orçado": new_orcado,
                "Observações": new_observacoes,
            }

            if not new_supplier_name:
                st.error("É preciso informar um nome para o fornecedor.")
            elif new_supplier_name in suppliers:
                st.error("Esse fornecedor já existe.")
            else:
                
                new_df = pd.DataFrame(columns=ALL_COLUMNS)
                new_df.loc[len(new_df)] = new_row
                st.session_state.suppliers_data[new_supplier_name] = new_df
                save_fornecedores()
                st.session_state.fornecedor_criado = True

    else:
        # Edição de Fornecedor existente
        if selected_supplier in suppliers:
            st.subheader(f"Edição do Fornecedor: {selected_supplier}")
            df_original = st.session_state.suppliers_data[selected_supplier].copy()
            if not df_original.empty:
                general_info = df_original.iloc[0][GENERAL_COLUMNS].to_dict()
            else:
                general_info = {col: "" for col in GENERAL_COLUMNS}

            for col in GENERAL_COLUMNS:
                general_info[col] = st.text_input(col, value=general_info[col], key=f"{selected_supplier}_{col}")

            st.subheader("Produtos/Serviços")
            st.caption("Para inserir ou excluir linhas, use o '+' ou a lixeira no st.data_editor.")

            if "Inicio do contrato" in df_original.columns:
                df_original["Inicio do contrato"] = df_original["Inicio do contrato"].apply(_datetime_to_str)
            if "Termino do contrato" in df_original.columns:
                df_original["Termino do contrato"] = df_original["Termino do contrato"].apply(_datetime_to_str)

            column_config = {c: st.column_config.Column(disabled=True) for c in GENERAL_COLUMNS}
            edited_df = st.data_editor(
                df_original,
                column_config=column_config,
                num_rows="dynamic",
                key=f"editor_{selected_supplier}"
            )

            if not edited_df.empty:
                for col in GENERAL_COLUMNS:
                    edited_df[col] = general_info[col]

            st.session_state.suppliers_data[selected_supplier] = edited_df

            col1, col2 = st.columns(2)
            with col1:
                if st.button("Salvar Edições de Fornecedor", key=f"salvar_{selected_supplier}"):
                    save_fornecedores()
            with col2:
                if st.button("Excluir Fornecedor", key=f"excluir_{selected_supplier}"):
                    st.session_state.suppliers_data.pop(selected_supplier)
                    save_fornecedores()
                    st.warning(f"Fornecedor '{selected_supplier}' excluído!")
                    st.stop()

###############################################################################
# ABA 2: LISTA DE FORNECEDORES
###############################################################################
with tabs[1]:
    st.title("Lista de Fornecedores")
    suppliers = list(st.session_state.suppliers_data.keys())
    if suppliers:
        df_combined = pd.DataFrame()
        for sup_name in suppliers:
            df_temp = st.session_state.suppliers_data[sup_name].copy()
            df_temp.insert(0, "Aba (Fornecedor)", sup_name)
            df_combined = pd.concat([df_combined, df_temp], ignore_index=True)
        st.dataframe(df_combined)
    else:
        st.info("Não há fornecedores cadastrados.")

###############################################################################
# ABA 3: REGISTRAR PAGAMENTOS
###############################################################################
with tabs[2]:
    st.title("Registrar Pagamentos")
    st.write("Nesta seção, você pode adicionar novos pagamentos e salvá-los diretamente no Excel do SharePoint.")

    df_mensal = st.session_state["controle_mensal"].copy()

    # Selecionar Fornecedor
    fornecedores_list = list(st.session_state.suppliers_data.keys())
    sel_fornecedor = st.selectbox("Selecione o Fornecedor", [""] + fornecedores_list)

    # ID Fornecedor (auto)
    id_fornecedor_label = ""
    if sel_fornecedor:
        df_temp_forn = st.session_state.suppliers_data[sel_fornecedor]
        if not df_temp_forn.empty and "ID - Fornecedor" in df_temp_forn.columns:
            id_fornecedor_label = df_temp_forn.iloc[0]["ID - Fornecedor"]
    st.text_input("ID - Fornecedor", value=id_fornecedor_label, disabled=True)

    # ID Pagamento
    list_id_pag = []
    if sel_fornecedor:
        df_temp_forn = st.session_state.suppliers_data[sel_fornecedor]
        if "ID - Pagamento" in df_temp_forn.columns:
            vals = df_temp_forn["ID - Pagamento"].dropna().unique()
            list_id_pag = sorted([v for v in vals if v])

    chosen_id_pag = st.selectbox("ID - Pagamento Existente", ["(Novo)"] + list_id_pag)
    if chosen_id_pag == "(Novo)":
        typed_id_pag = st.text_input("Ou digite novo ID - Pagamento")
        final_id_pag = typed_id_pag.strip()
    else:
        final_id_pag = chosen_id_pag

    sel_categoria = st.selectbox("Categoria do Produto/Serviço", category_options)
    dia_vencimento = st.text_input("Dia Vencimento")
    data_envio_str = st.text_input("Data Envio (DD/MM/AAAA)")
    data_pagamento_str = st.text_input("Data Pagamento (DD/MM/AAAA)")
    metodo_pagamento = st.selectbox("Método de Pagamento", ["CARTÃO", "BOLETO"])
    status_pag = st.selectbox("Status de Pagamento", STATUS_PAG_OPCOES)
    planejado = st.selectbox("Planejado", ["SIM", "NÃO"])
    sel_moeda = st.selectbox("Moeda", MOEDAS_COMUNS)
    val_estimado_str = st.text_input("Valor Estimado (R$)")
    val_pago_str = st.text_input("Valor Pago Convertido (R$)")
    obs = st.text_input("Observações")

    hoje = datetime.date.today()
    ano_padrao = str(hoje.year)
    mes_padrao = MESES_ORDENADOS[hoje.month - 1] if 1 <= hoje.month <= 12 else "JANEIRO"
    sel_ano = st.text_input("Ano", ano_padrao)
    sel_mes = st.selectbox("Mês", MESES_ORDENADOS, index=(hoje.month - 1 if 1 <= hoje.month <= 12 else 0))

    if st.button("Salvar Pagamento Agora"):
        if not sel_fornecedor:
            st.error("Selecione o fornecedor.")
        else:
            try:
                dt_env = parse_date_br(data_envio_str)
                dt_pag = parse_date_br(data_pagamento_str)
                val_est = parse_float_br(val_estimado_str) or 0.0
                val_pag = parse_float_br(val_pago_str) or 0.0
                dif = val_est - val_pag

                new_row = {
                    "Fornecedor": sel_fornecedor,
                    "ID - Fornecedor": id_fornecedor_label,
                    "ID - Pagamento": final_id_pag,
                    "Categoria": sel_categoria,
                    "Dia Vencimento": dia_vencimento,
                    "Data Envio": dt_env,
                    "Data Pagamento": dt_pag,
                    "Metodo de Pagamento": metodo_pagamento,
                    "Status de Pagamento": status_pag,
                    "Planejado": planejado,
                    "Moeda": sel_moeda,
                    "Valor Estimado - Real": val_est,
                    "Valor Pago Convertido": val_pag,
                    "Diferença": dif,
                    "Observações": obs,
                    "Ano": sel_ano,
                    "Mes": sel_mes,
                }

                
                new_line_df = pd.DataFrame([new_row])
                df_mensal = pd.concat([df_mensal, new_line_df], ignore_index=True)

                # Ajusta datas
                if "Data Envio" in df_mensal.columns:
                    df_mensal["Data Envio"] = pd.to_datetime(df_mensal["Data Envio"], errors="coerce")
                if "Data Pagamento" in df_mensal.columns:
                    df_mensal["Data Pagamento"] = pd.to_datetime(df_mensal["Data Pagamento"], errors="coerce")

                st.session_state["controle_mensal"] = df_mensal

                # Salva imediatamente no SharePoint
                save_controle_mensal()

                st.success("Pagamento registrado com sucesso no Excel do SharePoint!!")
            except Exception as e:
                st.error(f"Erro ao lançar pagamento: {e}")
