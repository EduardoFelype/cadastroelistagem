import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime, date
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# Configuração da página
st.set_page_config(
    page_title="Sistema de Ordens de Serviço - Painel Completo",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Função para conectar ao banco de dados
@st.cache_resource
def init_database():
    conn = sqlite3.connect('ordens_servico_completo.db', check_same_thread=False)
    cursor = conn.cursor()
    
    # Criar tabela com todas as colunas da planilha
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS ordens_servico (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descricao_operacao TEXT,
            numero_oportunidade TEXT,
            numero_vta TEXT,
            numero_cotacao TEXT,
            numero_circuito TEXT,
            status_cotacao TEXT,
            denominacao_produto TEXT,
            quantidade INTEGER,
            status TEXT,
            valor_pedido_bruto REAL,
            criado_em DATE,
            emissor_ordem TEXT,
            nome_emissor_ordem TEXT,
            nome_gerente_contas TEXT,
            organizacao_vendas TEXT,
            canal_distribuicao TEXT,
            setor_atividade TEXT,
            item_sd TEXT,
            id_produto TEXT,
            tempo_contrato TEXT,
            data_importacao DATETIME DEFAULT CURRENT_TIMESTAMP,
            data_atualizacao DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    conn.commit()
    return conn

# Função para mapear status
def mapear_status(status_excel):
    if pd.isna(status_excel) or status_excel == '':
        return 'Pendente'
    
    status = str(status_excel).strip()
    status_map = {
        'concluído': 'Concluído',
        'concluido': 'Concluído',
        'finalizado': 'Concluído',
        'completo': 'Concluído',
        'pendente': 'Pendente',
        'aberto': 'Aberto',
        'liberado': 'Liberado',
        'liberada': 'Liberado',
        'aprovado': 'Aprovado',
        'em andamento': 'Em Andamento',
        'processando': 'Em Andamento'
    }
    
    return status_map.get(status.lower(), status)

# Função para converter data
def converter_data(data_excel):
    if pd.isna(data_excel) or data_excel == '':
        return None
    
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
            return (pd.to_datetime('1900-01-01') + pd.Timedelta(days=data_excel-2)).date()
        
        return pd.to_datetime(data_excel).date()
    except:
        return None

# Função para carregar dados
@st.cache_data
def carregar_dados():
    conn = init_database()
    df = pd.read_sql_query("SELECT * FROM ordens_servico ORDER BY id DESC", conn)
    return df

# Função para limpar dados antigos
def limpar_dados_antigos():
    conn = init_database()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM ordens_servico")
    conn.commit()
    st.cache_data.clear()

# Função para processar Excel completo
def processar_excel_completo(arquivo_excel, atualizar_dados=True):
    try:
        # Ler Excel
        df = pd.read_excel(arquivo_excel)
        
        st.info(f"📊 Planilha carregada: {len(df)} linhas, {len(df.columns)} colunas")
        
        # Mapear colunas da planilha
        colunas_esperadas = {
            'Descrição d/operação': 'descricao_operacao',
            'Número da Oportunidade': 'numero_oportunidade', 
            'Número da VTA': 'numero_vta',
            'Número da Cotação': 'numero_cotacao',
            'Número do Circuito': 'numero_circuito',
            'Status cotação': 'status_cotacao',
            'Denominação produto': 'denominacao_produto',
            'Quantidade': 'quantidade',
            'Status': 'status',
            'Valor pedido bruto': 'valor_pedido_bruto',
            'Criado em': 'criado_em',
            'Emissor da Ordem': 'emissor_ordem',
            'Nome do Emissor da Ordem': 'nome_emissor_ordem',
            'Nome do Gerente de Contas': 'nome_gerente_contas',
            'Organização de Vendas': 'organizacao_vendas',
            'Canal de distribuição': 'canal_distribuicao',
            'Setor de atividade': 'setor_atividade',
            'Item (SD)': 'item_sd',
            'ID produto': 'id_produto',
            'Tempo de Contrato': 'tempo_contrato'
        }
        
        # Verificar colunas existentes
        colunas_encontradas = {}
        for col_excel, col_db in colunas_esperadas.items():
            if col_excel in df.columns:
                colunas_encontradas[col_db] = col_excel
        
        st.success(f"✅ Encontradas {len(colunas_encontradas)} colunas de {len(colunas_esperadas)} esperadas")
        
        # Limpar dados se for atualização
        if atualizar_dados:
            limpar_dados_antigos()
            st.info("🗑️ Dados antigos removidos para atualização")
        
        # Processar dados
        registros_importados = 0
        erros = []
        
        conn = init_database()
        cursor = conn.cursor()
        
        # Preparar query de inserção
        colunas_sql = ', '.join(colunas_encontradas.keys())
        placeholders = ', '.join(['?' for _ in colunas_encontradas])
        
        query = f'''
            INSERT INTO ordens_servico ({colunas_sql})
            VALUES ({placeholders})
        '''
        
        # Barra de progresso
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, row in df.iterrows():
            try:
                # Preparar valores
                valores = []
                for col_db, col_excel in colunas_encontradas.items():
                    valor = row[col_excel]
                    
                    # Tratamento especial para diferentes tipos de dados
                    if col_db == 'criado_em':
                        valor = converter_data(valor)
                    elif col_db == 'status':
                        valor = mapear_status(valor)
                    elif col_db == 'quantidade':
                        try:
                            valor = int(valor) if not pd.isna(valor) else 0
                        except:
                            valor = 0
                    elif col_db == 'valor_pedido_bruto':
                        try:
                            valor = float(valor) if not pd.isna(valor) else 0.0
                        except:
                            valor = 0.0
                    elif pd.isna(valor):
                        valor = None
                    else:
                        valor = str(valor).strip() if valor else None
                    
                    valores.append(valor)
                
                # Inserir no banco
                cursor.execute(query, valores)
                registros_importados += 1
                
                # Atualizar progresso
                if registros_importados % 50 == 0:
                    progress = registros_importados / len(df)
                    progress_bar.progress(progress)
                    status_text.text(f"Processando... {registros_importados}/{len(df)} registros")
                
            except Exception as e:
                erros.append(f"Linha {idx+2}: {str(e)}")
        
        conn.commit()
        progress_bar.progress(1.0)
        status_text.text(f"✅ Concluído! {registros_importados} registros importados")
        
        st.cache_data.clear()
        
        return True, f"Importação concluída! {registros_importados} registros importados. Erros: {len(erros)}"
        
    except Exception as e:
        return False, f"Erro ao processar arquivo: {str(e)}"

# Interface principal
def main():
    st.title("📊 Sistema de Ordens de Serviço - Painel Completo")
    st.markdown("*Sistema integrado com planilha CARGA_PAINEL.xlsx - Atualização semanal*")
    
    # Sidebar
    st.sidebar.title("📋 Menu Principal")
    opcao = st.sidebar.selectbox(
        "Escolha uma opção:",
        ["📊 Dashboard Executivo", "📁 Atualizar Planilha", "🔍 Consultar Dados", "📈 Relatórios", "⚙️ Configurações"]
    )
    
    # Dashboard Executivo
    if opcao == "📊 Dashboard Executivo":
        st.header("📊 Dashboard Executivo")
        
        df = carregar_dados()
        
        if not df.empty:
            # Métricas principais
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                st.metric("Total de Ordens", len(df))
            
            with col2:
                concluidas = len(df[df['status'] == 'Concluído'])
                st.metric("Concluídas", concluidas)
            
            with col3:
                pendentes = len(df[df['status'] == 'Pendente'])
                st.metric("Pendentes", pendentes)
            
            with col4:
                valor_total = df['valor_pedido_bruto'].sum()
                st.metric("Valor Total", f"R$ {valor_total:,.2f}")
            
            with col5:
                ultima_atualizacao = df['data_importacao'].max() if 'data_importacao' in df.columns else 'N/A'
                st.metric("Última Atualização", str(ultima_atualizacao)[:10] if ultima_atualizacao != 'N/A' else 'N/A')
            
            # Gráficos
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📊 Status das Ordens")
                status_counts = df['status'].value_counts()
                fig_pie = px.pie(
                    values=status_counts.values,
                    names=status_counts.index,
                    title="Distribuição por Status",
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                st.plotly_chart(fig_pie, use_container_width=True)
            
            with col2:
                st.subheader("💰 Valor por Status de Cotação")
                valor_por_status = df.groupby('status_cotacao')['valor_pedido_bruto'].sum().sort_values(ascending=False)
                fig_bar = px.bar(
                    x=valor_por_status.index,
                    y=valor_por_status.values,
                    title="Valor Total por Status de Cotação",
                    labels={'y': 'Valor (R$)', 'x': 'Status da Cotação'}
                )
                st.plotly_chart(fig_bar, use_container_width=True)
            
            # Timeline de criação
            if 'criado_em' in df.columns:
                st.subheader("📅 Timeline de Criação das Ordens")
                df_timeline = df.copy()
                df_timeline['criado_em'] = pd.to_datetime(df_timeline['criado_em'])
                df_timeline['mes_ano'] = df_timeline['criado_em'].dt.to_period('M').astype(str)
                
                timeline_data = df_timeline.groupby('mes_ano').size().reset_index(name='quantidade')
                
                fig_timeline = px.line(
                    timeline_data,
                    x='mes_ano',
                    y='quantidade',
                    title="Ordens Criadas por Mês",
                    markers=True
                )
                fig_timeline.update_xaxes(tickangle=45)
                st.plotly_chart(fig_timeline, use_container_width=True)
            
            # Top clientes/produtos
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("🏢 Top 10 Clientes")
                top_clientes = df['nome_emissor_ordem'].value_counts().head(10)
                st.dataframe(top_clientes, use_container_width=True)
            
            with col2:
                st.subheader("📦 Top 10 Produtos")
                top_produtos = df['denominacao_produto'].value_counts().head(10)
                st.dataframe(top_produtos, use_container_width=True)
        else:
            st.info("📋 Nenhum dado encontrado. Faça a importação da planilha primeiro.")
    
    # Atualizar Planilha
    elif opcao == "📁 Atualizar Planilha":
        st.header("📁 Atualização Semanal da Planilha")
        
        st.info("""
        🔄 **Processo de Atualização Semanal**
        
        1. Faça upload da planilha CARGA_PAINEL.xlsx atualizada
        2. O sistema irá substituir todos os dados existentes
        3. Todos os registros serão atualizados com a nova versão
        4. Recomendado: Fazer backup antes da atualização
        """)
        
        df_atual = carregar_dados()
        if not df_atual.empty:
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Registros Atuais", len(df_atual))
            with col2:
                ultima_atualizacao = df_atual['data_importacao'].max() if 'data_importacao' in df_atual.columns else 'N/A'
                st.metric("Última Atualização", str(ultima_atualizacao)[:10] if ultima_atualizacao != 'N/A' else 'N/A')
            with col3:
                st.metric("Status", "✅ Dados Carregados")
        
        arquivo_excel = st.file_uploader(
            "📎 Selecione a planilha CARGA_PAINEL.xlsx atualizada",
            type=['xlsx', 'xls'],
            help="Faça upload da planilha completa para atualização semanal"
        )
        
        if arquivo_excel is not None:
            st.success(f"📁 Arquivo carregado: {arquivo_excel.name}")
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("🔄 Atualizar Dados Completos", type="primary", use_container_width=True):
                    with st.spinner("🔄 Processando atualização completa..."):
                        sucesso, mensagem = processar_excel_completo(arquivo_excel, atualizar_dados=True)
                        
                        if sucesso:
                            st.success(mensagem)
                            st.balloons()
                        else:
                            st.error(mensagem)
            
            with col2:
                if st.button("➕ Adicionar aos Dados Existentes", type="secondary", use_container_width=True):
                    with st.spinner("➕ Adicionando novos dados..."):
                        sucesso, mensagem = processar_excel_completo(arquivo_excel, atualizar_dados=False)
                        
                        if sucesso:
                            st.success(mensagem)
                        else:
                            st.error(mensagem)
    
    # Consultar Dados
    elif opcao == "🔍 Consultar Dados":
        st.header("🔍 Consulta Detalhada de Dados")
        
        df = carregar_dados()
        
        if not df.empty:
            # Filtros avançados
            st.subheader("🎛️ Filtros")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                filtro_status = st.multiselect(
                    "Status",
                    options=df['status'].unique(),
                    default=df['status'].unique()
                )
            
            with col2:
                filtro_status_cotacao = st.multiselect(
                    "Status Cotação",
                    options=df['status_cotacao'].unique() if 'status_cotacao' in df.columns else [],
                    default=df['status_cotacao'].unique() if 'status_cotacao' in df.columns else []
                )
            
            with col3:
                filtro_produto = st.selectbox(
                    "Produto",
                    options=['Todos'] + list(df['denominacao_produto'].unique()) if 'denominacao_produto' in df.columns else ['Todos'],
                    index=0
                )
            
            with col4:
                filtro_cliente = st.text_input("🔍 Buscar Cliente")
            
            # Aplicar filtros
            df_filtrado = df[df['status'].isin(filtro_status)]
            
            if 'status_cotacao' in df.columns and filtro_status_cotacao:
                df_filtrado = df_filtrado[df_filtrado['status_cotacao'].isin(filtro_status_cotacao)]
            
            if filtro_produto != 'Todos':
                df_filtrado = df_filtrado[df_filtrado['denominacao_produto'] == filtro_produto]
            
            if filtro_cliente:
                df_filtrado = df_filtrado[
                    df_filtrado['nome_emissor_ordem'].str.contains(filtro_cliente, case=False, na=False) |
                    df_filtrado['descricao_operacao'].str.contains(filtro_cliente, case=False, na=False)
                ]
            
            # Exibir resultados
            st.subheader(f"📊 Resultados: {len(df_filtrado)} registros")
            
            # Colunas para exibir
            colunas_exibir = [
                'descricao_operacao', 'denominacao_produto', 'status', 'status_cotacao',
                'valor_pedido_bruto', 'criado_em', 'nome_emissor_ordem'
            ]
            colunas_disponiveis = [col for col in colunas_exibir if col in df_filtrado.columns]
            
            st.dataframe(
                df_filtrado[colunas_disponiveis],
                use_container_width=True,
                height=400
            )
            
            # Download dos dados filtrados
            if st.button("📥 Download Dados Filtrados (Excel)"):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_filtrado.to_excel(writer, index=False, sheet_name='Dados_Filtrados')
                
                st.download_button(
                    label="📥 Baixar Excel",
                    data=output.getvalue(),
                    file_name=f"dados_filtrados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.info("📋 Nenhum dado encontrado.")
    
    # Relatórios
    elif opcao == "📈 Relatórios":
        st.header("📈 Relatórios Gerenciais")
        
        df = carregar_dados()
        
        if not df.empty:
            # Relatório por período
            st.subheader("📅 Relatório por Período")
            
            if 'criado_em' in df.columns:
                df_periodo = df.copy()
                df_periodo['criado_em'] = pd.to_datetime(df_periodo['criado_em'])
                
                col1, col2 = st.columns(2)
                with col1:
                    data_inicio = st.date_input("Data Início", value=df_periodo['criado_em'].min().date())
                with col2:
                    data_fim = st.date_input("Data Fim", value=df_periodo['criado_em'].max().date())
                
                # Filtrar por período
                mask = (df_periodo['criado_em'].dt.date >= data_inicio) & (df_periodo['criado_em'].dt.date <= data_fim)
                df_periodo_filtrado = df_periodo.loc[mask]
                
                if not df_periodo_filtrado.empty:
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Ordens no Período", len(df_periodo_filtrado))
                    with col2:
                        valor_periodo = df_periodo_filtrado['valor_pedido_bruto'].sum()
                        st.metric("Valor Total", f"R$ {valor_periodo:,.2f}")
                    with col3:
                        ticket_medio = df_periodo_filtrado['valor_pedido_bruto'].mean()
                        st.metric("Ticket Médio", f"R$ {ticket_medio:,.2f}")
                    
                    # Gráfico de evolução
                    df_evolucao = df_periodo_filtrado.groupby(df_periodo_filtrado['criado_em'].dt.date).agg({
                        'id': 'count',
                        'valor_pedido_bruto': 'sum'
                    }).reset_index()
                    
                    fig_evolucao = px.line(
                        df_evolucao,
                        x='criado_em',
                        y='id',
                        title="Evolução de Ordens no Período",
                        labels={'id': 'Quantidade de Ordens', 'criado_em': 'Data'}
                    )
                    st.plotly_chart(fig_evolucao, use_container_width=True)
            
            # Relatório de performance
            st.subheader("🎯 Performance por Status")
            
            performance = df.groupby(['status', 'status_cotacao']).agg({
                'id': 'count',
                'valor_pedido_bruto': ['sum', 'mean']
            }).round(2)
            
            st.dataframe(performance, use_container_width=True)
        else:
            st.info("📋 Nenhum dado para relatórios.")
    
    # Configurações
    elif opcao == "⚙️ Configurações":
        st.header("⚙️ Configurações do Sistema")
        
        df = carregar_dados()
        
        st.subheader("📊 Informações do Banco de Dados")
        if not df.empty:
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total de Registros", len(df))
            with col2:
                tamanho_mb = len(df) * 0.001  # Estimativa
                st.metric("Tamanho Estimado", f"{tamanho_mb:.2f} MB")
            with col3:
                colunas_count = len(df.columns)
                st.metric("Colunas", colunas_count)
        
        st.subheader("🗑️ Limpeza de Dados")
        st.warning("⚠️ Atenção: Esta ação é irreversível!")
        
        if st.button("🗑️ Limpar Todos os Dados", type="secondary"):
            if st.session_state.get('confirmar_limpeza'):
                limpar_dados_antigos()
                st.success("✅ Todos os dados foram removidos!")
                st.rerun()
            else:
                st.session_state.confirmar_limpeza = True
                st.warning("⚠️ Clique novamente para confirmar a limpeza completa.")
        
        st.subheader("📋 Estrutura da Planilha Esperada")
        colunas_esperadas = [
            "Descrição d/operação", "Número da Oportunidade", "Número da VTA",
            "Número da Cotação", "Número do Circuito", "Status cotação",
            "Denominação produto", "Quantidade", "Status", "Valor pedido bruto",
            "Criado em", "Emissor da Ordem", "Nome do Emissor da Ordem",
            "Nome do Gerente de Contas", "Organização de Vendas",
            "Canal de distribuição", "Setor de atividade", "Item (SD)",
            "ID produto", "Tempo de Contrato"
        ]
        
        st.info("📋 Colunas esperadas na planilha CARGA_PAINEL.xlsx:")
        for i, col in enumerate(colunas_esperadas, 1):
            st.text(f"{i:2d}. {col}")

if __name__ == "__main__":
    main()