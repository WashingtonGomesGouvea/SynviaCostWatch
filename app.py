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
    "CNPJ",            # Mantemos como string
    "Contato",         # Também string (padrão telefone)
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
    "Valor mensal",
    "Valor do plano",
    "Observações",
    "Forma de pagamento",
    "Tempo de pagamento",
]

ALL_COLUMNS = GENERAL_COLUMNS + SPECIFIC_COLUMNS

# -------------------------------------------------------------
# 3) FUNÇÃO PARA CONVERTER STRING DE VALOR -> FLOAT
# -------------------------------------------------------------
def parse_float_br(value):
    """
    Converte strings no formato 'R$ 5.400,00' ou '5.400,00' em float (5400.00).
    Ajuste conforme seu Excel.
    """
    if isinstance(value, str):
        value = value.replace("R$", "").replace(".", "").replace(",", ".").strip()
        try:
            return float(value)
        except ValueError:
            return None
    return value

# -------------------------------------------------------------
# 3.1) FUNÇÕES PARA GERAR ID AUTOMÁTICO
# -------------------------------------------------------------
def generate_id_fornecedor(supplier_name):
    """
    Gera um ID automático, pegando as 3 primeiras letras do nome (limpas) + número aleatório.
    Ex: 'Synvia' -> 'SYN123'
    """
    if not supplier_name:
        return ""
    prefix = re.sub(r"[^A-Za-z]", "", supplier_name).upper()[:3]
    rand_num = random.randint(100, 999)
    return f"{prefix}{rand_num}"

def generate_id_produto(produto_name):
    """
    Gera um ID automático para o produto, começando com 'P' + 3 letras do nome + número aleatório.
    Ex: 'Internet Link' -> 'PINT456'
    """
    if not produto_name:
        return ""
    prefix = re.sub(r"[^A-Za-z]", "", produto_name).upper()[:3]
    rand_num = random.randint(100, 999)
    return f"P{prefix}{rand_num}"

# -------------------------------------------------------------
# 4) CARREGAR DADOS DO SHAREPOINT
# -------------------------------------------------------------
def load_data():
    """
    Lê o arquivo Excel do SharePoint e retorna um dicionário {aba: DataFrame}.
    Ignora abas 'MATRIZ' e 'MODELO'.
    Converte 'CNPJ', 'Contato' e colunas de data para string para evitar problemas no st.data_editor.
    """
    try:
        ctx = ClientContext(SITE_URL).with_credentials(
            UserCredential(EMAIL_REMETENTE, SENHA_EMAIL)
        )
        response = File.open_binary(ctx, FILE_URL)
        excel_data = response.content  # bytes

        sheets = pd.read_excel(BytesIO(excel_data), sheet_name=None)

        # Ignora abas "MATRIZ" e "MODELO"
        filtered_sheets = {
            name: df for name, df in sheets.items()
            if name not in ["MATRIZ", "MODELO"]
        }

        for sheet_name, df in filtered_sheets.items():
            # Garante todas as colunas
            for col in ALL_COLUMNS:
                if col not in df.columns:
                    df[col] = ""

            # Converte 'CNPJ' e 'Contato' em string
            if "CNPJ" in df.columns:
                df["CNPJ"] = df["CNPJ"].astype(str)
            if "Contato" in df.columns:
                df["Contato"] = df["Contato"].astype(str)

            # Se "Inicio do contrato" / "Termino do contrato" vier como datetime, converte para string "DD/MM/AAAA"
            if "Inicio do contrato" in df.columns:
                df["Inicio do contrato"] = df["Inicio do contrato"].apply(_datetime_to_str)
            if "Termino do contrato" in df.columns:
                df["Termino do contrato"] = df["Termino do contrato"].apply(_datetime_to_str)

            # Converte valores monetários
            if "Valor mensal" in df.columns:
                df["Valor mensal"] = df["Valor mensal"].apply(parse_float_br)
            if "Valor do plano" in df.columns:
                df["Valor do plano"] = df["Valor do plano"].apply(parse_float_br)

            filtered_sheets[sheet_name] = df

        return filtered_sheets

    except Exception as e:
        st.error(f"Erro ao carregar o arquivo Excel: {e}")
        return {}

def _datetime_to_str(val):
    """
    Converte um valor datetime para string no formato DD/MM/AAAA.
    Se for NaT, retorna "".
    Se já for string, retorna como está.
    """
    if pd.isnull(val):
        return ""
    if isinstance(val, (pd.Timestamp, datetime.datetime)):
        return val.strftime("%d/%m/%Y")
    return str(val)  # Se já for string, ou outro tipo

# -------------------------------------------------------------
# 5) SALVAR DADOS DE VOLTA NO SHAREPOINT
# -------------------------------------------------------------
def save_data():
    """
    Converte datas para texto (dd/mm/aaaa) antes de salvar no Excel.
    Garante que 'CNPJ' e 'Contato' permaneçam como string.
    """
    try:
        for supplier_name, df in st.session_state.suppliers_data.items():
            # Garante que 'CNPJ' e 'Contato' são string
            if "CNPJ" in df.columns:
                df["CNPJ"] = df["CNPJ"].astype(str)
            if "Contato" in df.columns:
                df["Contato"] = df["Contato"].astype(str)

            # Converte datas em texto
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
            st.warning(
                "O arquivo está bloqueado para edição (aberto por outra pessoa ou em uso). "
                "Feche o arquivo ou faça check-in para liberá-lo antes de salvar."
            )
        else:
            st.error(f"Erro ao salvar o arquivo Excel: {e}")

# -------------------------------------------------------------
# 6) INICIALIZA O ST.SESSION_STATE
# -------------------------------------------------------------
if "suppliers_data" not in st.session_state:
    st.session_state.suppliers_data = load_data()

suppliers = list(st.session_state.suppliers_data.keys())

# -------------------------------------------------------------
# 7) ABA(S) NO STREAMLIT
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

    if selected_supplier == "Adicionar Novo Fornecedor":
        st.subheader("Adicionar Novo Fornecedor")

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

        new_cnpj = st.text_input("CNPJ")  # string
        new_contato = st.text_input("Contato (ex: (12) 982896323)")  # string
        new_centro_custo = st.text_input("Centro de custo")

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

        def_tempo = ""
        if new_forma_pagamento == "A Vista":
            def_tempo = "1"
        new_tempo_pagamento = st.text_input("Tempo de pagamento (Parcelas)", value=def_tempo)

        new_inicio_str = st.text_input("Início do contrato (DD/MM/AAAA)", "")
        new_termino_str = st.text_input("Término do contrato (DD/MM/AAAA)", "")
        new_valor_mensal_str = st.text_input("Valor Mensal (R$)", "")
        new_valor_plano_str = st.text_input("Valor do Plano (R$)", "")

        if st.button("Criar Fornecedor"):
            if not new_supplier_name:
                st.error("É preciso informar um nome para o fornecedor.")
            elif new_supplier_name in suppliers:
                st.error("Esse fornecedor já existe.")
            else:
                new_df = pd.DataFrame(columns=ALL_COLUMNS)

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
                    "Tempo de pagamento": new_tempo_pagamento,
                }

                # Converte datas para texto dd/mm/aaaa
                try:
                    dt_inicio = datetime.datetime.strptime(new_inicio_str, "%d/%m/%Y") if new_inicio_str else None
                except ValueError:
                    dt_inicio = None
                if dt_inicio:
                    new_row["Inicio do contrato"] = dt_inicio.strftime("%d/%m/%Y")
                else:
                    new_row["Inicio do contrato"] = ""

                try:
                    dt_termino = datetime.datetime.strptime(new_termino_str, "%d/%m/%Y") if new_termino_str else None
                except ValueError:
                    dt_termino = None
                if dt_termino:
                    new_row["Termino do contrato"] = dt_termino.strftime("%d/%m/%Y")
                else:
                    new_row["Termino do contrato"] = ""

                new_row["Valor mensal"] = parse_float_br(new_valor_mensal_str)
                new_row["Valor do plano"] = parse_float_br(new_valor_plano_str)

                new_df.loc[len(new_df)] = new_row

                st.session_state.suppliers_data[new_supplier_name] = new_df
                st.success(f"Fornecedor '{new_supplier_name}' criado com sucesso!")

                # Limpa session state
                st.session_state.supplier_name = ""
                st.session_state.product_name = ""
                st.session_state.auto_id_fornecedor = ""
                st.session_state.auto_id_produto = ""

                save_data()
                suppliers = list(st.session_state.suppliers_data.keys())

    else:
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

        # Converte as colunas de data para string no df_original, se ainda estiverem como datetime
        if "Inicio do contrato" in df_original.columns:
            df_original["Inicio do contrato"] = df_original["Inicio do contrato"].apply(_datetime_to_str)
        if "Termino do contrato" in df_original.columns:
            df_original["Termino do contrato"] = df_original["Termino do contrato"].apply(_datetime_to_str)

        # Garante que CNPJ e Contato sejam string
        if "CNPJ" in df_original.columns:
            df_original["CNPJ"] = df_original["CNPJ"].astype(str)
        if "Contato" in df_original.columns:
            df_original["Contato"] = df_original["Contato"].astype(str)

        # Configura colunas gerais como disabled no data_editor
        column_config = {col: st.column_config.Column(disabled=True) for col in GENERAL_COLUMNS}

        # Use TextColumn para datas e CNPJ/Contato
        if "Inicio do contrato" in df_original.columns:
            column_config["Inicio do contrato"] = st.column_config.TextColumn("Início do Contrato")
        if "Termino do contrato" in df_original.columns:
            column_config["Termino do contrato"] = st.column_config.TextColumn("Término do Contrato")
        if "CNPJ" in df_original.columns:
            column_config["CNPJ"] = st.column_config.TextColumn("CNPJ")
        if "Contato" in df_original.columns:
            column_config["Contato"] = st.column_config.TextColumn("Contato")

        # Mantemos NumberColumn para valores
        if "Valor mensal" in df_original.columns:
            column_config["Valor mensal"] = st.column_config.NumberColumn(
                "Valor Mensal",
                format="%.2f"
            )
        if "Valor do plano" in df_original.columns:
            column_config["Valor do plano"] = st.column_config.NumberColumn(
                "Valor do Plano",
                format="%.2f"
            )

        edited_df = st.data_editor(
            df_original,
            column_config=column_config,
            num_rows="dynamic",
            key=f"{supplier}_editor"
        )

        # Sincroniza as colunas gerais com text_input
        for col in GENERAL_COLUMNS:
            edited_df[col] = general_info[col]

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
                st.warning("Por favor, recarregue a página para atualizar a lista.")
                st.stop()

# -------------------------------------------------------------
# ABA 2: LISTA DE FORNECEDORES
# -------------------------------------------------------------
with tabs[1]:
    st.title("Lista de Fornecedores")

    # Link para Power BI
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

    # Link para pasta com arquivos-fonte
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
