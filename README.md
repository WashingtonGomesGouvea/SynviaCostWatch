# SynviaCostWatch 🚀

Este projeto implementa uma aplicação **Streamlit** para gerenciar dados de **fornecedores** armazenados em um arquivo **Excel** hospedado no **SharePoint**. A aplicação permite:

- ✨ **Criar** novos fornecedores (criando novas abas no Excel).  
- 📝 **Editar** dados de fornecedores existentes.  
- ❌ **Excluir** fornecedores (removendo a aba correspondente do Excel).  
- ➕➖ **Inserir** e **remover** linhas dentro de cada fornecedor, usando o `st.data_editor`.  
- 📊 **Listar** todos os fornecedores em uma única tabela, **unificando** as abas.

## Funcionalidades

1. **Interface amigável** com duas abas:
   - **Gerenciar Fornecedores**: Criação, edição e exclusão.
   - **Lista de Fornecedores**: Visualização de todos os fornecedores e colunas em uma só tabela.
2. **Criação de abas** automaticamente no Excel ao salvar.
3. **Inserção e remoção de linhas** de forma dinâmica com `st.data_editor`.
4. **Integração** com SharePoint via `office365-rest-python-client`.
5. **Armazenamento seguro** de credenciais em `secrets.toml`.

## Requisitos

- **Python 3.8+**
- **Bibliotecas Python** (listadas no arquivo `requirements.txt`)

### Dependências

As principais dependências estão no arquivo `requirements.txt`, por exemplo:


Para instalar todas as dependências, execute:

```bash
pip install -r requirements.txt


[sharepoint]
email = "email.teste@email.com"
password = "SUA_SENHA_DE_APLICATIVO"
site_url = "https://seusite.sharepoint.com/sites/"
file_url = "/sites/gestaodeprodutos/Documentos Compartilhados/Gestão financeira/Controle dos Fornecedores - AutomationTest.xlsx"

streamlit run app.py

SynviaCostWatch/
 ├─ .streamlit/
 │   └─ secrets.toml              # Armazena credenciais e URLs do SharePoint
 ├─ app.py                        # Código principal em Streamlit
 ├─ requirements.txt              # Dependências
 └─ README.md                     # Documentação do projeto
 ```
## Como Usar:

### Gerenciar Fornecedores

Adicionar novo fornecedor (gera nova aba no Excel).
Editar informações gerais (nome, CNPJ, contato etc.).
Inserir/remover linhas de produtos/serviços.
Excluir fornecedor, removendo a aba correspondente.

## Lista de Fornecedores

Visualiza todas as abas (fornecedores) em uma única tabela.
Permite filtrar e analisar dados de forma centralizada.

## Observações
Se o arquivo no SharePoint estiver aberto por outra pessoa, você pode receber um erro 423 Locked. Nesse caso, feche o arquivo ou faça check-in antes de salvar.
Se “Documentos Compartilhados” não funcionar, tente “Shared Documents” (depende do nome interno da biblioteca).
Mantenha as credenciais fora do repositório público, usando .gitignore.

## Como Contribuir
Contribuições são bem-vindas! Sinta-se à vontade para abrir um issue ou enviar um pull request com melhorias, correções ou novas funcionalidades. 💖

## Licença
Este projeto não possui uma licença específica definida. Caso deseje, adicione um arquivo LICENSE para definir os termos de uso e distribuição.

Feito com 💼 + ☕ por Washington Gomes