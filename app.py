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
    "CNPJ",
    "Contato",
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
    return f"P{prefix}{rand_num}"   # <-- agora inicia com "P"

# -------------------------------------------------------------
# 4) CARREGAR DADOS DO SHAREPOINT
# -------------------------------------------------------------
def load_data():
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

        # Ajusta cada aba
        for sheet_name, df in filtered_sheets.items():
            for col in ALL_COLUMNS:
                if col not in df.columns:
                    df[col] = ""

            # Se estiver em texto no Excel, tudo bem; se estiver em datetime,
            # não vamos converter, pois agora queremos strings.
            # (Mas se quiser converter datas antigas para string, pode fazer aqui.)
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
# 5) SALVAR DADOS DE VOLTA NO SHAREPOINT
# -------------------------------------------------------------
def save_data():
    try:
        # Para cada fornecedor (aba), transformamos datas em texto
        for supplier_name, df in st.session_state.suppliers_data.items():

            if "Inicio do contrato" in df.columns:
                for i in range(len(df)):
                    val = df.at[i, "Inicio do contrato"]
                    # Se val for datetime ou Timestamp válido, converte; se for NaT, salva como ""
                    if pd.isnull(val):  # Significa que é NaT ou None
                        df.at[i, "Inicio do contrato"] = ""
                    elif isinstance(val, (pd.Timestamp, datetime.datetime)):
                        df.at[i, "Inicio do contrato"] = val.strftime("%d/%m/%Y")

            if "Termino do contrato" in df.columns:
                for i in range(len(df)):
                    val = df.at[i, "Termino do contrato"]
                    if pd.isnull(val):
                        df.at[i, "Termino do contrato"] = ""
                    elif isinstance(val, (pd.Timestamp, datetime.datetime)):
                        df.at[i, "Termino do contrato"] = val.strftime("%d/%m/%Y")

        # Agora gravamos no Excel
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

    # 7.1) CRIAÇÃO DE NOVO FORNECEDOR
    if selected_supplier == "Adicionar Novo Fornecedor":
        st.subheader("Adicionar Novo Fornecedor")

        # Para não regenerar ID toda hora, guardamos nome do fornecedor e produto no session_state
        if "supplier_name" not in st.session_state:
            st.session_state.supplier_name = ""
        if "product_name" not in st.session_state:
            st.session_state.product_name = ""

        # Nome do fornecedor
        new_supplier_name = st.text_input("Nome do Fornecedor", value=st.session_state.supplier_name)

        # Se o nome mudou, gera ID novamente
        if new_supplier_name != st.session_state.supplier_name:
            st.session_state.supplier_name = new_supplier_name
            st.session_state.auto_id_fornecedor = generate_id_fornecedor(new_supplier_name)

        # Exibe ID gerado (mas permite edição manual)
        new_id_fornecedor = st.text_input("ID - Fornecedor (auto)",
                                          value=st.session_state.get("auto_id_fornecedor", ""))

        new_cnpj = st.text_input("CNPJ")
        new_contato = st.text_input("Contato")
        new_centro_custo = st.text_input("Centro de custo")

        # Descrição do Produto e ID do Produto
        new_descricao_produto = st.text_input("Descrição do Produto", value=st.session_state.product_name)

        # Se a descrição mudar, gera ID novamente
        if new_descricao_produto != st.session_state.product_name:
            st.session_state.product_name = new_descricao_produto
            st.session_state.auto_id_produto = generate_id_produto(new_descricao_produto)

        new_id_produto = st.text_input("ID - Produto (auto)",
                                       value=st.session_state.get("auto_id_produto", ""))

        # Status fixo
        new_status = "ATIVO"

        # Localidade com opções
        localidades = ["PAULINIA", "AMBAS", "CAMPINAS"]
        new_localidade = st.selectbox("Localidade", localidades)

        # Metodo de pagamento
        new_metodo_pagamento = st.selectbox("Método de Pagamento", ["BOLETO", "CARTÃO"])

        # Forma de pagamento
        forma_options = ["A Prazo", "A Vista"]
        new_forma_pagamento = st.selectbox("Forma de pagamento", forma_options)

        # Se for "A Vista" e o usuário não digitar nada, define como "1"
        def_tempo = ""
        if new_forma_pagamento == "A Vista":
            def_tempo = "1"
        new_tempo_pagamento = st.text_input("Tempo de pagamento (Parcelas)", value=def_tempo)

        # Datas e valores
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
                # Cria DataFrame vazio
                new_df = pd.DataFrame(columns=ALL_COLUMNS)

                # Monta dicionário
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

                # Limpa session state para não regenerar IDs na próxima criação
                st.session_state.supplier_name = ""
                st.session_state.product_name = ""
                st.session_state.auto_id_fornecedor = ""
                st.session_state.auto_id_produto = ""

                save_data()
                suppliers = list(st.session_state.suppliers_data.keys())

    # 7.2) EDIÇÃO DE FORNECEDOR EXISTENTE
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

        # Configura colunas gerais como disabled no data_editor
        column_config = {col: st.column_config.Column(disabled=True) for col in GENERAL_COLUMNS}

        # Mantemos a configuração de colunas para data e valores
        if "Inicio do contrato" in df_original.columns:
            column_config["Inicio do contrato"] = st.column_config.DateColumn(
                "Início do Contrato",
                format="DD/MM/YYYY"
            )
        if "Termino do contrato" in df_original.columns:
            column_config["Termino do contrato"] = st.column_config.DateColumn(
                "Término do Contrato",
                format="DD/MM/YYYY"
            )
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

        # Converte datas do editor para string (dd/mm/aaaa) ao final
        # (Faremos isso em save_data(), mas se quiser forçar aqui, pode.)
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
