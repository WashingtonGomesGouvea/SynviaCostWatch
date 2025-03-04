import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
import random
import re

# Biblioteca para conexão com SharePoint
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.user_credential import UserCredential

# -------------------------------------------------------------
# 1) CREDENCIAIS E CAMINHO DO SHAREPOINT (via st.secrets)
# -------------------------------------------------------------
EMAIL_REMETENTE = st.secrets["sharepoint"]["email"]
SENHA_EMAIL = st.secrets["sharepoint"]["password"]
SITE_URL = st.secrets["sharepoint"]["site_url"]
FILE_URL = st.secrets["sharepoint"]["file_url"]

# -------------------------------------------------------------
# 2) DEFINIÇÃO DAS COLUNAS
# -------------------------------------------------------------
GENERAL_COLUMNS = [
    "Fornecedor",
    "ID - Fornecedor",
    "CNPJ",      # String
    "Contato",   # String
    "Centro de custo",
]

SPECIFIC_COLUMNS = [
    "Nº do Serviço",
    "ID - Produto",
    "Descrição do Produto",
    "Localidade",
    "Status",
    "Inicio do contrato",
    "Termino do contrato",
    "Tempo de contrato",
    "Metodo de pagamento",
    "Tipo de pagamento",
    "Dia de Pagamento",
    "ID - Pagamento",
    "Status de Pagamento",
    "Orçado",
    "Valor mensal",       # Campo que o usuário digita
    "Valor do plano",     # Auto-calculado
    "Observações",
    "Forma de pagamento",
    "Tempo de pagamento", # Parcelas
]

ALL_COLUMNS = GENERAL_COLUMNS + SPECIFIC_COLUMNS

# -------------------------------------------------------------
# 3) FUNÇÕES AUXILIARES
# -------------------------------------------------------------
def parse_float_br(value):
    """
    Converte strings como:
      - "R$ 1.234,56"  -> 1234.56
      - "10,0"         -> 10.0
      - "10.0"         -> 10.0
      - "1.234.567,89" -> 1234567.89
    Sem remover o ponto decimal se for o único ponto.
    """
    if not isinstance(value, str):
        return value

    # Remove "R$" e espaços
    value = value.replace("R$", "").strip()

    # Se contiver ponto e vírgula, assumimos que ponto é milhar e vírgula é decimal
    if "." in value and "," in value:
        # Remove pontos
        value = value.replace(".", "")
        # Troca vírgula por ponto
        value = value.replace(",", ".")
    elif "," in value and "." not in value:
        # Se só tiver vírgula, trocamos por ponto
        value = value.replace(",", ".")
    # Se só tiver ponto, consideramos decimal
    # Se não tiver ponto nem vírgula, segue como está

    # Tenta converter para float
    try:
        return float(value)
    except ValueError:
        return None

def _datetime_to_str(val):
    """
    Converte datetime/NaT em string DD/MM/AAAA ou retorna str(val) se já for string.
    """
    if pd.isnull(val):
        return ""
    if isinstance(val, (pd.Timestamp, datetime.datetime)):
        return val.strftime("%d/%m/%Y")
    return str(val)

def generate_id_fornecedor(supplier_name):
    """Ex: 'Synvia' -> 'SYN123'."""
    if not supplier_name:
        return ""
    prefix = re.sub(r"[^A-Za-z]", "", supplier_name).upper()[:3]
    rand_num = random.randint(100, 999)
    return f"{prefix}{rand_num}"

def generate_id_produto(produto_name):
    """Ex: 'Internet Link' -> 'PINT456'."""
    if not produto_name:
        return ""
    prefix = re.sub(r"[^A-Za-z]", "", produto_name).upper()[:3]
    rand_num = random.randint(100, 999)
    return f"P{prefix}{rand_num}"

# -------------------------------------------------------------
# 4) CARREGAR DADOS DO SHAREPOINT
# -------------------------------------------------------------
def load_data():
    """Lê Excel do SharePoint e retorna {aba: DataFrame}, ignorando 'MATRIZ' e 'MODELO'."""
    try:
        ctx = ClientContext(SITE_URL).with_credentials(
            UserCredential(EMAIL_REMETENTE, SENHA_EMAIL)
        )
        response = File.open_binary(ctx, FILE_URL)
        excel_data = response.content

        sheets = pd.read_excel(BytesIO(excel_data), sheet_name=None)
        filtered_sheets = {
            name: df for name, df in sheets.items() if name not in ["MATRIZ", "MODELO"]
        }

        for sheet_name, df in filtered_sheets.items():
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

            if "Valor mensal" in df.columns:
                df["Valor mensal"] = df["Valor mensal"].apply(parse_float_br)
            if "Valor do plano" in df.columns:
                df["Valor do plano"] = df["Valor do plano"].apply(parse_float_br)

            filtered_sheets[sheet_name] = df

        return filtered_sheets

    except Exception as e:
        st.error(f"Erro ao carregar o arquivo Excel: {e}")
        return {}

# -------------------------------------------------------------
# 5) SALVAR DADOS
# -------------------------------------------------------------
def save_data():
    """Converte datas para texto e salva no SharePoint."""
    try:
        for supplier_name, df in st.session_state.suppliers_data.items():
            if "CNPJ" in df.columns:
                df["CNPJ"] = df["CNPJ"].astype(str)
            if "Contato" in df.columns:
                df["Contato"] = df["Contato"].astype(str)

            if "Inicio do contrato" in df.columns:
                df["Inicio do contrato"] = df["Inicio do contrato"].apply(_datetime_to_str)
            if "Termino do contrato" in df.columns:
                df["Termino do contrato"] = df["Termino do contrato"].apply(_datetime_to_str)

        output = BytesIO()
        with pd.ExcelWriter(output) as writer:
            for supplier_name, df in st.session_state.suppliers_data.items():
                df.to_excel(writer, sheet_name=supplier_name, index=False)
        output.seek(0)

        ctx = ClientContext(SITE_URL).with_credentials(
            UserCredential(EMAIL_REMETENTE, SENHA_EMAIL)
        )
        File.save_binary(ctx, FILE_URL, output.read())

        st.success("Dados salvos com sucesso!")
    except Exception as e:
        if "Locked" in str(e) or "423" in str(e):
            st.warning("O arquivo está bloqueado. Feche ou faça check-in antes de salvar.")
        else:
            st.error(f"Erro ao salvar o arquivo Excel: {e}")

# -------------------------------------------------------------
# 6) AUTO-CÁLCULO
# -------------------------------------------------------------
def auto_calc_valor_plano():
    """Callback para atualizar Valor do plano."""
    val_mensal_str = st.session_state.get("new_valor_mensal_str", "")
    parcelas_str = st.session_state.get("new_tempo_pagamento", "")

    val_mensal = parse_float_br(val_mensal_str)
    try:
        parcelas = int(parcelas_str)
    except ValueError:
        parcelas = 1

    if val_mensal is not None and parcelas > 0:
        total = val_mensal * parcelas
        st.session_state["new_valor_plano_str"] = f"{total:.2f}"
    else:
        st.session_state["new_valor_plano_str"] = ""

# -------------------------------------------------------------
# 7) INICIALIZA ST.SESSION_STATE
# -------------------------------------------------------------
if "suppliers_data" not in st.session_state:
    st.session_state.suppliers_data = load_data()

if "fornecedor_criado" not in st.session_state:
    st.session_state.fornecedor_criado = False

for key in ["new_valor_mensal_str", "new_tempo_pagamento", "new_valor_plano_str"]:
    if key not in st.session_state:
        st.session_state[key] = ""

suppliers = list(st.session_state.suppliers_data.keys())

# -------------------------------------------------------------
# 8) ABA(S) NO STREAMLIT
# -------------------------------------------------------------
tabs = st.tabs(["Gerenciar Fornecedores", "Lista de Fornecedores"])

# -------------------------------------------------------------
# ABA 1: GERENCIAR FORNECEDORES
# -------------------------------------------------------------
with tabs[0]:
    st.title("Gestão de Fornecedores")

    selected_supplier = st.sidebar.selectbox(
        "Selecione o Fornecedor",
        ["Adicionar Novo Fornecedor"] + suppliers
    )

    # Se o fornecedor foi criado, mostramos mensagem e um botão para criar outro
    if st.session_state.fornecedor_criado and selected_supplier == "Adicionar Novo Fornecedor":
        st.success("Fornecedor criado com sucesso!")
        if st.button("Criar outro fornecedor"):
            # Limpa manualmente as variáveis
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
        # Não exibe o formulário novamente neste ciclo
    elif selected_supplier == "Adicionar Novo Fornecedor" and not st.session_state.fornecedor_criado:
        st.subheader("Adicionar Novo Fornecedor")

        # IDs
        if "supplier_name" not in st.session_state:
            st.session_state.supplier_name = ""
        if "product_name" not in st.session_state:
            st.session_state.product_name = ""

        new_supplier_name = st.text_input("Nome do Fornecedor", value=st.session_state.supplier_name)
        if new_supplier_name != st.session_state.supplier_name:
            st.session_state.supplier_name = new_supplier_name
            st.session_state.auto_id_fornecedor = generate_id_fornecedor(new_supplier_name)

        new_id_fornecedor = st.text_input(
            "ID - Fornecedor (auto)",
            value=st.session_state.get("auto_id_fornecedor", "")
        )

        new_cnpj = st.text_input("CNPJ")
        new_contato = st.text_input("Contato")
        new_centro_custo = st.text_input("Centro de custo")

        # Produto
        new_descricao_produto = st.text_input("Descrição do Produto", value=st.session_state.product_name)
        if new_descricao_produto != st.session_state.product_name:
            st.session_state.product_name = new_descricao_produto
            st.session_state.auto_id_produto = generate_id_produto(new_descricao_produto)

        new_id_produto = st.text_input(
            "ID - Produto (auto)",
            value=st.session_state.get("auto_id_produto", "")
        )

        new_status = "ATIVO"
        localidades = ["PAULINIA", "AMBAS", "CAMPINAS"]
        new_localidade = st.selectbox("Localidade", localidades)

        new_metodo_pagamento = st.selectbox("Método de Pagamento", ["BOLETO", "CARTÃO"])

        forma_options = ["A Prazo", "A Vista"]
        new_forma_pagamento = st.selectbox("Forma de pagamento", forma_options)

        # Se for "A Vista", define parcelas = 1
        if new_forma_pagamento == "A Vista":
            st.session_state["new_tempo_pagamento"] = "1"

        # Valor Mensal e Tempo de pagamento (callbacks)
        st.text_input(
            "Valor Mensal (R$)",
            key="new_valor_mensal_str",
            on_change=auto_calc_valor_plano
        )
        st.text_input(
            "Tempo de pagamento (Parcelas)",
            key="new_tempo_pagamento",
            on_change=auto_calc_valor_plano
        )
        st.text_input(
            "Valor do Plano (R$) - Autopreenchido",
            key="new_valor_plano_str",
            disabled=True
        )

        # Datas
        new_inicio_str = st.text_input("Início do contrato (DD/MM/AAAA)", "")
        new_termino_str = st.text_input("Término do contrato (DD/MM/AAAA)", "")

        if st.button("Criar Fornecedor"):
            val_mensal = parse_float_br(st.session_state["new_valor_mensal_str"])
            val_plano = parse_float_br(st.session_state["new_valor_plano_str"])

            new_df = pd.DataFrame(columns=ALL_COLUMNS)

            try:
                dt_inicio = datetime.datetime.strptime(new_inicio_str, "%d/%m/%Y") if new_inicio_str else None
            except ValueError:
                dt_inicio = None
            if dt_inicio:
                inicio_fmt = dt_inicio.strftime("%d/%m/%Y")
            else:
                inicio_fmt = ""

            try:
                dt_termino = datetime.datetime.strptime(new_termino_str, "%d/%m/%Y") if new_termino_str else None
            except ValueError:
                dt_termino = None
            if dt_termino:
                termino_fmt = dt_termino.strftime("%d/%m/%Y")
            else:
                termino_fmt = ""

            new_row = {
                "Fornecedor": new_supplier_name,
                "ID - Fornecedor": new_id_fornecedor,
                "CNPJ": new_cnpj,
                "Contato": new_contato,
                "Centro de custo": new_centro_custo,
                "Descrição do Produto": new_descricao_produto,
                "ID - Produto": new_id_produto,
                "Status": new_status,
                "Localidade": new_localidade,
                "Metodo de pagamento": new_metodo_pagamento,
                "Forma de pagamento": new_forma_pagamento,
                "Tempo de pagamento": st.session_state["new_tempo_pagamento"],
                "Valor mensal": val_mensal,
                "Valor do plano": val_plano,
                "Inicio do contrato": inicio_fmt,
                "Termino do contrato": termino_fmt
            }

            if not new_supplier_name:
                st.error("É preciso informar um nome para o fornecedor.")
            elif new_supplier_name in suppliers:
                st.error("Esse fornecedor já existe.")
            else:
                new_df.loc[len(new_df)] = new_row
                st.session_state.suppliers_data[new_supplier_name] = new_df
                save_data()

                # Marca que foi criado e não exibe o form de novo
                st.session_state.fornecedor_criado = True

    else:
        # EDIÇÃO DE FORNECEDOR EXISTENTE
        supplier = selected_supplier
        df_original = st.session_state.suppliers_data[supplier].copy()

        if not df_original.empty:
            general_info = df_original.iloc[0][GENERAL_COLUMNS].to_dict()
        else:
            general_info = {col: "" for col in GENERAL_COLUMNS}

        st.subheader(f"Informações Gerais - {supplier}")
        for col in GENERAL_COLUMNS:
            general_info[col] = st.text_input(
                col,
                value=general_info[col],
                key=f"{supplier}_{col}"
            )

        st.subheader("Produtos/Serviços")
        st.caption("Para inserir ou excluir linhas, use o '+' ou a lixeira no `st.data_editor`.")

        if "Inicio do contrato" in df_original.columns:
            df_original["Inicio do contrato"] = df_original["Inicio do contrato"].apply(_datetime_to_str)
        if "Termino do contrato" in df_original.columns:
            df_original["Termino do contrato"] = df_original["Termino do contrato"].apply(_datetime_to_str)
        if "CNPJ" in df_original.columns:
            df_original["CNPJ"] = df_original["CNPJ"].astype(str)
        if "Contato" in df_original.columns:
            df_original["Contato"] = df_original["Contato"].astype(str)

        column_config = {col: st.column_config.Column(disabled=True) for col in GENERAL_COLUMNS}

        if "Inicio do contrato" in df_original.columns:
            column_config["Inicio do contrato"] = st.column_config.TextColumn("Início do Contrato")
        if "Termino do contrato" in df_original.columns:
            column_config["Termino do contrato"] = st.column_config.TextColumn("Término do Contrato")
        if "CNPJ" in df_original.columns:
            column_config["CNPJ"] = st.column_config.TextColumn("CNPJ")
        if "Contato" in df_original.columns:
            column_config["Contato"] = st.column_config.TextColumn("Contato")

        if "Valor mensal" in df_original.columns:
            column_config["Valor mensal"] = st.column_config.NumberColumn("Valor Mensal", format="%.2f")
        if "Valor do plano" in df_original.columns:
            column_config["Valor do plano"] = st.column_config.NumberColumn("Valor do Plano", format="%.2f")

        edited_df = st.data_editor(
            df_original,
            column_config=column_config,
            num_rows="dynamic",
            key=f"{supplier}_editor"
        )

        if not edited_df.empty:
            for col in GENERAL_COLUMNS:
                edited_df.at[0, col] = general_info[col]

        st.session_state.suppliers_data[supplier] = edited_df

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Salvar", key=f"{supplier}_save"):
                save_data()

        with col2:
            if st.button("Excluir Fornecedor", key=f"{supplier}_delete"):
                st.session_state.suppliers_data.pop(supplier)
                st.success(f"Fornecedor '{supplier}' excluído com sucesso!")
                save_data()
                # Mensagem de aviso para recarregar a página
                st.warning("Por favor, recarregue a página para atualizar a lista.")
                st.stop()

# -------------------------------------------------------------
# ABA 2: LISTA DE FORNECEDORES
# -------------------------------------------------------------
with tabs[1]:
    st.title("Lista de Fornecedores")

    st.write("### Dashboard no Power BI")
    st.markdown(
        "[Clique aqui para visualizar o relatório Power BI](https://app.powerbi.com/reportEmbed?reportId=cf2f800d-cf4a-4cb7-b871-99e583f70aa8&autoAuth=true&ctid=fee1b506-24b6-444a-919e-83df9442dc5d)",
        unsafe_allow_html=True
    )
    st.markdown(
        """
        <iframe title="Análise de Custos" width="1140" height="541.25"
                src="https://app.powerbi.com/reportEmbed?reportId=cf2f800d-cf4a-4cb7-b871-99e583f70aa8&autoAuth=true&ctid=fee1b506-24b6-444a-919e-83df9442dc5d"
                frameborder="0" allowFullScreen="true"></iframe>
        """,
        unsafe_allow_html=True
    )

    st.write("### Pasta SharePoint com arquivos-fonte")
    st.markdown(
        "[Acesse aqui a pasta no SharePoint](https://synviagroup.sharepoint.com/:f:/r/sites/gestaodeprodutos/Documentos%20Compartilhados/Gest%C3%A3o%20financeira?csf=1&web=1&e=vHmqxV)",
        unsafe_allow_html=True
    )

    if suppliers:
        st.write("Exibindo todas as colunas do Excel, unificando os dados de cada aba.")
        combined_df = pd.DataFrame()
        for sup_name in suppliers:
            df_sup = st.session_state.suppliers_data[sup_name].copy()
            df_sup.insert(0, "Aba", sup_name)
            combined_df = pd.concat([combined_df, df_sup], ignore_index=True)

        st.dataframe(combined_df)
    else:
        st.write("Não há fornecedores cadastrados.")
