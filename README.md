Sistema de Ordens de Serviço
Um sistema simples para gerenciamento de ordens de serviço, desenvolvido com Streamlit, SQLite e Pandas. Permite cadastrar, listar, editar, deletar ordens e importar dados a partir de planilhas Excel.

Funcionalidades
Dashboard com métricas e gráficos de ordens por situação e prioridade.

Cadastro manual de ordens de serviço.

Listagem com filtros por situação, prioridade e cliente.

Importação de ordens via arquivo Excel.

Edição e exclusão de ordens cadastradas.

Persistência de dados usando SQLite.

Tecnologias
Python 3.8+

Streamlit

Pandas

Plotly

SQLite

Requisitos
Instale as dependências via pip:


pip install streamlit pandas plotly openpyxl
Como usar
Clone o repositório ou copie os arquivos.

Instale as dependências.

Execute o app:


streamlit run app.py
Navegue pelo menu lateral para usar as funcionalidades.

Importação de Excel
O arquivo deve conter colunas como:

Descrição d/operação → Nome do Cliente

Denominação produto → Descrição do Serviço

Criado em → Data de Abertura

Status → Situação (ex: concluído, pendente)

Observações
O banco SQLite é criado localmente no arquivo ordens_servico.db.

Ao importar Excel, é possível optar por limpar dados existentes ou ignorar duplicatas.

Prioridade padrão das importações é "média".

Contato
Qualquer dúvida ou sugestão, entre em contato!
