import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime, date
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# Configuração da página
st.set_page_config(
    page_title="Sistema de Ordens de Serviço",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Função para conectar ao banco de dados
@st.cache_resource
def init_database():
    conn = sqlite3.connect('ordens_servico.db', check_same_thread=False)
    cursor = conn.cursor()
    
    # Criar tabela se não existir
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS ordens_servico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_cliente TEXT NOT NULL,
            descricao_servico TEXT NOT NULL,
            data_abertura DATE NOT NULL,
            prioridade TEXT NOT NULL CHECK (prioridade IN ('baixa', 'media', 'alta')),
            situacao TEXT NOT NULL CHECK (situacao IN ('aberto', 'em_andamento', 'concluido')),
            data_criacao DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    conn.commit()
    return conn

# Função para mapear status do Excel
def mapear_status(status_excel):
    status_map = {
        'concluído': 'concluido',
        'concluido': 'concluido',
        'finalizado': 'concluido',
        'completo': 'concluido',
        'pendente': 'aberto',
        'aberto': 'aberto',
        'novo': 'aberto',
        'iniciado': 'aberto',
        'em andamento': 'em_andamento',
        'em_andamento': 'em_andamento',
        'processando': 'em_andamento',
        'executando': 'em_andamento'
    }
    return status_map.get(status_excel.lower().strip(), 'aberto')

# Função para converter data
def converter_data(data_excel):
    if pd.isna(data_excel) or data_excel == '':
        return date.today()
    
    try:
        if isinstance(data_excel, str):
            # Tentar diferentes formatos de data
            for fmt in ['%d.%m.%Y %H:%M:%S', '%d/%m/%Y', '%Y-%m-%d', '%d.%m.%Y']:
                try:
                    return datetime.strptime(data_excel, fmt).date()
                except:
                    continue
        elif isinstance(data_excel, (int, float)):
            # Data do Excel (número de dias desde 1900-01-01)
            return pd.to_datetime('1900-01-01') + pd.Timedelta(days=data_excel-2)
        
        return pd.to_datetime(data_excel).date()
    except:
        return date.today()

# Função para carregar dados
@st.cache_data
def carregar_dados():
    conn = init_database()
    df = pd.read_sql_query("SELECT * FROM ordens_servico ORDER BY id DESC", conn)
    return df

# Função para inserir ordem
def inserir_ordem(nome_cliente, descricao_servico, data_abertura, prioridade, situacao):
    conn = init_database()
    cursor = conn.cursor()
    
    cursor.execute('''
        INSERT INTO ordens_servico (nome_cliente, descricao_servico, data_abertura, prioridade, situacao)
        VALUES (?, ?, ?, ?, ?)
    ''', (nome_cliente, descricao_servico, data_abertura, prioridade, situacao))
    
    conn.commit()
    st.cache_data.clear()

# Função para atualizar ordem
def atualizar_ordem(id_ordem, nome_cliente, descricao_servico, data_abertura, prioridade, situacao):
    conn = init_database()
    cursor = conn.cursor()
    
    cursor.execute('''
        UPDATE ordens_servico 
        SET nome_cliente=?, descricao_servico=?, data_abertura=?, prioridade=?, situacao=?
        WHERE id=?
    ''', (nome_cliente, descricao_servico, data_abertura, prioridade, situacao, id_ordem))
    
    conn.commit()
    st.cache_data.clear()

# Função para deletar ordem
def deletar_ordem(id_ordem):
    conn = init_database()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM ordens_servico WHERE id=?", (id_ordem,))
    conn.commit()
    st.cache_data.clear()

# Função para processar Excel
def processar_excel(arquivo_excel, limpar_dados=False, ignorar_duplicatas=True):
    try:
        # Ler Excel
        df = pd.read_excel(arquivo_excel)
        
        # Mapear colunas
        colunas_mapeadas = {}
        for col in df.columns:
            col_lower = str(col).lower().strip()
            if 'descrição d/operação' in col_lower or 'cliente' in col_lower:
                colunas_mapeadas['nome_cliente'] = col
            elif 'denominação produto' in col_lower or 'descrição' in col_lower:
                colunas_mapeadas['descricao_servico'] = col
            elif 'criado em' in col_lower or 'data' in col_lower:
                colunas_mapeadas['data_abertura'] = col
            elif 'status' in col_lower and 'cotação' not in col_lower:
                colunas_mapeadas['situacao'] = col
        
        # Verificar se colunas essenciais foram encontradas
        if not all(k in colunas_mapeadas for k in ['nome_cliente', 'descricao_servico', 'situacao']):
            return False, "Colunas obrigatórias não encontradas na planilha"
        
        # Limpar dados existentes se solicitado
        if limpar_dados:
            conn = init_database()
            cursor = conn.cursor()
            cursor.execute("DELETE FROM ordens_servico")
            conn.commit()
        
        # Processar dados
        registros_importados = 0
        erros = []
        
        conn = init_database()
        cursor = conn.cursor()
        
        for idx, row in df.iterrows():
            try:
                nome_cliente = str(row[colunas_mapeadas['nome_cliente']]).strip()
                descricao_servico = str(row[colunas_mapeadas['descricao_servico']]).strip()
                
                if not nome_cliente or not descricao_servico or nome_cliente == 'nan' or descricao_servico == 'nan':
                    continue
                
                # Verificar duplicatas
                if ignorar_duplicatas:
                    cursor.execute(
                        "SELECT COUNT(*) FROM ordens_servico WHERE nome_cliente=? AND descricao_servico=?",
                        (nome_cliente, descricao_servico)
                    )
                    if cursor.fetchone()[0] > 0:
                        continue
                
                # Processar dados
                data_abertura = converter_data(row.get(colunas_mapeadas.get('data_abertura', ''), ''))
                situacao = mapear_status(str(row[colunas_mapeadas['situacao']]))
                prioridade = 'media'
                
                # Inserir no banco
                cursor.execute('''
                    INSERT INTO ordens_servico (nome_cliente, descricao_servico, data_abertura, prioridade, situacao)
                    VALUES (?, ?, ?, ?, ?)
                ''', (nome_cliente, descricao_servico, data_abertura, prioridade, situacao))
                
                registros_importados += 1
                
            except Exception as e:
                erros.append(f"Linha {idx+2}: {str(e)}")
        
        conn.commit()
        st.cache_data.clear()
        
        return True, f"Importação concluída! {registros_importados} registros importados. Erros: {len(erros)}"
        
    except Exception as e:
        return False, f"Erro ao processar arquivo: {str(e)}"

# Interface principal
def main():
    st.title("🔧 Sistema de Ordens de Serviço")
    
    # Sidebar
    st.sidebar.title("Menu")
    opcao = st.sidebar.selectbox(
        "Escolha uma opção:",
        ["📊 Dashboard", "➕ Cadastrar Ordem", "📋 Listar Ordens", "📁 Importar Excel", "✏️ Editar Ordem"]
    )
    
    # Dashboard
    if opcao == "📊 Dashboard":
        st.header("📊 Dashboard")
        
        df = carregar_dados()
        
        if not df.empty:
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total de Ordens", len(df))
            
            with col2:
                abertas = len(df[df['situacao'] == 'aberto'])
                st.metric("Abertas", abertas)
            
            with col3:
                em_andamento = len(df[df['situacao'] == 'em_andamento'])
                st.metric("Em Andamento", em_andamento)
            
            with col4:
                concluidas = len(df[df['situacao'] == 'concluido'])
                st.metric("Concluídas", concluidas)
            
            # Gráficos
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Distribuição por Situação")
                situacao_counts = df['situacao'].value_counts()
                fig_pie = px.pie(
                    values=situacao_counts.values,
                    names=situacao_counts.index,
                    title="Ordens por Situação"
                )
                st.plotly_chart(fig_pie, use_container_width=True)
            
            with col2:
                st.subheader("Distribuição por Prioridade")
                prioridade_counts = df['prioridade'].value_counts()
                fig_bar = px.bar(
                    x=prioridade_counts.index,
                    y=prioridade_counts.values,
                    title="Ordens por Prioridade"
                )
                st.plotly_chart(fig_bar, use_container_width=True)
            
            # Ordens recentes
            st.subheader("Ordens Recentes")
            st.dataframe(df.head(10), use_container_width=True)
        else:
            st.info("Nenhuma ordem cadastrada ainda.")
    
    # Cadastrar Ordem
    elif opcao == "➕ Cadastrar Ordem":
        st.header("➕ Cadastrar Nova Ordem")
        
        with st.form("form_cadastro"):
            nome_cliente = st.text_input("Nome do Cliente *")
            descricao_servico = st.text_area("Descrição do Serviço *")
            data_abertura = st.date_input("Data de Abertura", value=date.today())
            prioridade = st.selectbox("Prioridade", ["baixa", "media", "alta"])
            situacao = st.selectbox("Situação", ["aberto", "em_andamento", "concluido"])
            
            submitted = st.form_submit_button("Cadastrar Ordem")
            
            if submitted:
                if nome_cliente and descricao_servico:
                    inserir_ordem(nome_cliente, descricao_servico, data_abertura, prioridade, situacao)
                    st.success("Ordem cadastrada com sucesso!")
                    st.rerun()
                else:
                    st.error("Por favor, preencha todos os campos obrigatórios.")
    
    # Listar Ordens
    elif opcao == "📋 Listar Ordens":
        st.header("📋 Lista de Ordens de Serviço")
        
        df = carregar_dados()
        
        if not df.empty:
            # Filtros
            col1, col2, col3 = st.columns(3)
            
            with col1:
                filtro_situacao = st.multiselect(
                    "Filtrar por Situação",
                    options=df['situacao'].unique(),
                    default=df['situacao'].unique()
                )
            
            with col2:
                filtro_prioridade = st.multiselect(
                    "Filtrar por Prioridade",
                    options=df['prioridade'].unique(),
                    default=df['prioridade'].unique()
                )
            
            with col3:
                filtro_cliente = st.text_input("Buscar por Cliente")
            
            # Aplicar filtros
            df_filtrado = df[
                (df['situacao'].isin(filtro_situacao)) &
                (df['prioridade'].isin(filtro_prioridade))
            ]
            
            if filtro_cliente:
                df_filtrado = df_filtrado[
                    df_filtrado['nome_cliente'].str.contains(filtro_cliente, case=False, na=False)
                ]
            
            # Exibir dados
            st.dataframe(
                df_filtrado[['id', 'nome_cliente', 'descricao_servico', 'data_abertura', 'prioridade', 'situacao']],
                use_container_width=True
            )
            
            # Opção de deletar
            if st.checkbox("Modo de Exclusão"):
                id_para_deletar = st.number_input("ID da ordem para deletar", min_value=1, step=1)
                if st.button("Deletar Ordem", type="secondary"):
                    if st.session_state.get('confirmar_delete'):
                        deletar_ordem(id_para_deletar)
                        st.success("Ordem deletada com sucesso!")
                        st.rerun()
                    else:
                        st.session_state.confirmar_delete = True
                        st.warning("Clique novamente para confirmar a exclusão.")
        else:
            st.info("Nenhuma ordem cadastrada.")
    
    # Importar Excel
    elif opcao == "📁 Importar Excel":
        st.header("📁 Importar Planilha Excel")
        
        st.info("""
        **Formato esperado da planilha:**
        - **Descrição d/operação** → Nome do Cliente
        - **Denominação produto** → Descrição do Serviço  
        - **Criado em** → Data de Abertura
        - **Status** → Situação (Concluído → concluido, Pendente → aberto)
        """)
        
        arquivo_excel = st.file_uploader(
            "Selecione o arquivo Excel",
            type=['xlsx', 'xls'],
            help="Formatos aceitos: .xlsx, .xls"
        )
        
        col1, col2 = st.columns(2)
        with col1:
            limpar_dados = st.checkbox("Limpar dados existentes antes da importação")
        with col2:
            ignorar_duplicatas = st.checkbox("Ignorar registros duplicados", value=True)
        
        if arquivo_excel is not None:
            if st.button("Importar Planilha", type="primary"):
                with st.spinner("Processando arquivo..."):
                    sucesso, mensagem = processar_excel(arquivo_excel, limpar_dados, ignorar_duplicatas)
                    
                    if sucesso:
                        st.success(mensagem)
                    else:
                        st.error(mensagem)
    
    # Editar Ordem
    elif opcao == "✏️ Editar Ordem":
        st.header("✏️ Editar Ordem de Serviço")
        
        df = carregar_dados()
        
        if not df.empty:
            # Seletor de ordem
            opcoes_ordem = [f"#{row['id']} - {row['nome_cliente']}" for _, row in df.iterrows()]
            ordem_selecionada = st.selectbox("Selecione a ordem para editar:", opcoes_ordem)
            
            if ordem_selecionada:
                id_ordem = int(ordem_selecionada.split('#')[1].split(' -')[0])
                ordem_atual = df[df['id'] == id_ordem].iloc[0]
                
                with st.form("form_edicao"):
                    nome_cliente = st.text_input("Nome do Cliente", value=ordem_atual['nome_cliente'])
                    descricao_servico = st.text_area("Descrição do Serviço", value=ordem_atual['descricao_servico'])
                    data_abertura = st.date_input("Data de Abertura", value=pd.to_datetime(ordem_atual['data_abertura']).date())
                    prioridade = st.selectbox("Prioridade", ["baixa", "media", "alta"], index=["baixa", "media", "alta"].index(ordem_atual['prioridade']))
                    situacao = st.selectbox("Situação", ["aberto", "em_andamento", "concluido"], index=["aberto", "em_andamento", "concluido"].index(ordem_atual['situacao']))
                    
                    submitted = st.form_submit_button("Salvar Alterações")
                    
                    if submitted:
                        if nome_cliente and descricao_servico:
                            atualizar_ordem(id_ordem, nome_cliente, descricao_servico, data_abertura, prioridade, situacao)
                            st.success("Ordem atualizada com sucesso!")
                            st.rerun()
                        else:
                            st.error("Por favor, preencha todos os campos obrigatórios.")
        else:
            st.info("Nenhuma ordem cadastrada para editar.")

if __name__ == "__main__":
    main()