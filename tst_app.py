import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import re

# Configuração da página
st.set_page_config(
    page_title="Análise de Quantidades - Empréstimos Hospitalares",
    page_icon="📊",
    layout="wide"
)

# Título principal
st.title("📊 Análise de Quantidades - Empréstimos entre Hospitais")
st.markdown("---")

def extrair_numeros(documento):
    """Extrai a primeira sequência numérica encontrada no documento."""
    match = re.search(r'\d+', str(documento))
    return match.group() if match else ""

def _parse_date_column(series):
    """Converte a coluna de datas para datetime."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_datetime(series, unit='d', origin='1899-12-30', errors='coerce')
    return pd.to_datetime(series, dayfirst=True, errors='coerce')

@st.cache_data
def carregar_dados(arquivo_concedido, arquivo_recebido):
    """Carrega e processa os arquivos de empréstimos."""
    try:
        # Carregar arquivos
        df_saida = pd.read_excel(arquivo_concedido, engine='openpyxl')
        df_entrada = pd.read_excel(arquivo_recebido, engine='openpyxl')
        
        # Padronizar nomes das colunas
        df_saida.columns = [c.strip().lower() for c in df_saida.columns]
        df_entrada.columns = [c.strip().lower() for c in df_entrada.columns]
        
        # Converter datas
        df_saida['data'] = _parse_date_column(df_saida['data'])
        df_entrada['data'] = _parse_date_column(df_entrada['data'])
        
        # Garantir que documento seja string
        df_saida['documento'] = df_saida['documento'].astype(str).str.strip()
        df_entrada['documento'] = df_entrada['documento'].astype(str).str.strip()
        
        return df_saida, df_entrada
    except Exception as e:
        st.error(f"Erro ao carregar arquivos: {e}")
        return None, None

def analisar_quantidades(df_saida, df_entrada):
    """Analisa as quantidades de saída e entrada."""
    
    # Renomear colunas
    df_saida_temp = df_saida.rename(columns={
        'qt_entrada': 'Qtd Saída',
        'ds_produto': 'Produto (Saída)',
        'data': 'Data'
    })
    
    df_entrada_temp = df_entrada.rename(columns={
        'qt_entrada': 'Qtd Entrada',
        'ds_produto': 'Produto (Entrada)',
        'data': 'Data'
    })
    
    # Agrupar quantidades
    df_saida_grouped = df_saida_temp.groupby(
        ['Data', 'unidade_origem', 'unidade_destino', 'documento', 'Produto (Saída)'],
        dropna=False
    )['Qtd Saída'].sum().reset_index()
    
    df_entrada_grouped = df_entrada_temp.groupby(
        ['Data', 'unidade_origem', 'unidade_destino', 'documento', 'Produto (Entrada)'],
        dropna=False
    )['Qtd Entrada'].sum().reset_index()
    
    # Merge dos dados
    df_merged = pd.merge(
        df_saida_grouped,
        df_entrada_grouped,
        on=['documento', 'unidade_origem', 'unidade_destino'],
        how='left',
        suffixes=('_saida', '_entrada')
    )
    
    # Preencher NaN com 0
    df_merged['Qtd Entrada'] = df_merged['Qtd Entrada'].fillna(0)
    
    # Determinar coluna de data
    data_col = 'Data_saida' if 'Data_saida' in df_merged.columns else 'Data'
    
    # Agrupar novamente
    df_merged_final = df_merged.groupby(
        [data_col, 'unidade_origem', 'unidade_destino', 'documento', 'Produto (Saída)'],
        dropna=False
    ).agg(
        Qtd_Saida=('Qtd Saída', 'sum'),
        Qtd_Entrada=('Qtd Entrada', 'sum')
    ).reset_index()
    
    # Calcular diferença
    df_merged_final['Diferença Qtd'] = df_merged_final['Qtd_Saida'] - df_merged_final['Qtd_Entrada']
    
    # Renomear colunas para layout final
    df_resultado = df_merged_final.rename(columns={
        data_col: 'Data',
        'Qtd_Saida': 'Qtd Saída',
        'Qtd_Entrada': 'Qtd Entrada',
        'unidade_origem': 'Unidade Origem',
        'unidade_destino': 'Unidade Destino',
        'documento': 'Documento'
    })
    
    # Reordenar colunas
    df_resultado = df_resultado[[
        'Data', 'Produto (Saída)', 'Unidade Origem', 'Unidade Destino',
        'Qtd Saída', 'Qtd Entrada', 'Diferença Qtd', 'Documento'
    ]]
    
    return df_resultado

# Sidebar para upload de arquivos
st.sidebar.header("📁 Upload de Arquivos")
arquivo_concedido = st.sidebar.file_uploader(
    "Empréstimos Concedidos (Saída)",
    type=['xlsx'],
    help="Arquivo com os empréstimos concedidos"
)
arquivo_recebido = st.sidebar.file_uploader(
    "Empréstimos Recebidos (Entrada)",
    type=['xlsx'],
    help="Arquivo com os empréstimos recebidos"
)

# Processar dados se ambos os arquivos foram carregados
if arquivo_concedido and arquivo_recebido:
    df_saida, df_entrada = carregar_dados(arquivo_concedido, arquivo_recebido)
    
    if df_saida is not None and df_entrada is not None:
        # Analisar quantidades
        df_resultado = analisar_quantidades(df_saida, df_entrada)
        
        # Filtros na sidebar
        st.sidebar.markdown("---")
        st.sidebar.header("🔍 Filtros")
        
        # Filtro de data
        if not df_resultado['Data'].isna().all():
            data_min = df_resultado['Data'].min()
            data_max = df_resultado['Data'].max()
            
            data_inicio = st.sidebar.date_input(
                "Data Início",
                value=data_min,
                min_value=data_min,
                max_value=data_max
            )
            data_fim = st.sidebar.date_input(
                "Data Fim",
                value=data_max,
                min_value=data_min,
                max_value=data_max
            )
            
            # Aplicar filtro de data
            df_resultado = df_resultado[
                (df_resultado['Data'] >= pd.to_datetime(data_inicio)) &
                (df_resultado['Data'] <= pd.to_datetime(data_fim))
            ]
        
        # Filtro de unidade origem
        unidades_origem = ['Todas'] + sorted(df_resultado['Unidade Origem'].unique().tolist())
        unidade_origem_selecionada = st.sidebar.selectbox(
            "Unidade Origem",
            unidades_origem
        )
        
        if unidade_origem_selecionada != 'Todas':
            df_resultado = df_resultado[df_resultado['Unidade Origem'] == unidade_origem_selecionada]
        
        # Filtro de unidade destino
        unidades_destino = ['Todas'] + sorted(df_resultado['Unidade Destino'].unique().tolist())
        unidade_destino_selecionada = st.sidebar.selectbox(
            "Unidade Destino",
            unidades_destino
        )
        
        if unidade_destino_selecionada != 'Todas':
            df_resultado = df_resultado[df_resultado['Unidade Destino'] == unidade_destino_selecionada]
        
        # Filtro de divergência
        tipo_divergencia = st.sidebar.radio(
            "Tipo de Divergência",
            ['Todas', 'Apenas com Divergência (≠ 0)', 'Sem Recebimento (Entrada = 0)', 'Conformes (= 0)']
        )
        
        if tipo_divergencia == 'Apenas com Divergência (≠ 0)':
            df_resultado = df_resultado[df_resultado['Diferença Qtd'] != 0]
        elif tipo_divergencia == 'Sem Recebimento (Entrada = 0)':
            df_resultado = df_resultado[df_resultado['Qtd Entrada'] == 0]
        elif tipo_divergencia == 'Conformes (= 0)':
            df_resultado = df_resultado[df_resultado['Diferença Qtd'] == 0]
        
        # Métricas principais
        st.header("📈 Resumo Geral")
        col1, col2, col3, col4 = st.columns(4)
        
        total_registros = len(df_resultado)
        total_divergencias = len(df_resultado[df_resultado['Diferença Qtd'] != 0])
        total_sem_recebimento = len(df_resultado[df_resultado['Qtd Entrada'] == 0])
        total_conformes = len(df_resultado[df_resultado['Diferença Qtd'] == 0])
        
        col1.metric("Total de Registros", total_registros)
        col2.metric("Com Divergência", total_divergencias, 
                   delta=f"{(total_divergencias/total_registros*100):.1f}%" if total_registros > 0 else "0%")
        col3.metric("Sem Recebimento", total_sem_recebimento,
                   delta=f"{(total_sem_recebimento/total_registros*100):.1f}%" if total_registros > 0 else "0%")
        col4.metric("Conformes", total_conformes,
                   delta=f"{(total_conformes/total_registros*100):.1f}%" if total_registros > 0 else "0%")
        
        st.markdown("---")
        
        # Gráficos
        st.header("📊 Visualizações")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Gráfico de pizza - Status das divergências
            status_counts = pd.DataFrame({
                'Status': ['Conformes', 'Com Divergência', 'Sem Recebimento'],
                'Quantidade': [total_conformes, total_divergencias - total_sem_recebimento, total_sem_recebimento]
            })
            
            fig_pizza = px.pie(
                status_counts,
                values='Quantidade',
                names='Status',
                title='Distribuição por Status',
                color='Status',
                color_discrete_map={
                    'Conformes': '#00CC96',
                    'Com Divergência': '#FFA15A',
                    'Sem Recebimento': '#EF553B'
                }
            )
            st.plotly_chart(fig_pizza, use_container_width=True)
        
        with col2:
            # Top 10 produtos com maior divergência
            df_top_divergencias = df_resultado.nlargest(10, 'Diferença Qtd')[['Produto (Saída)', 'Diferença Qtd']]
            df_top_divergencias['Produto (Saída)'] = df_top_divergencias['Produto (Saída)'].str[:40] + '...'
            
            fig_bar = px.bar(
                df_top_divergencias,
                x='Diferença Qtd',
                y='Produto (Saída)',
                orientation='h',
                title='Top 10 Produtos com Maior Divergência',
                labels={'Diferença Qtd': 'Diferença de Quantidade', 'Produto (Saída)': 'Produto'},
                color='Diferença Qtd',
                color_continuous_scale='Reds'
            )
            fig_bar.update_layout(yaxis={'categoryorder': 'total ascending'})
            st.plotly_chart(fig_bar, use_container_width=True)
        
        # Gráfico de linha - Divergências ao longo do tempo
        if not df_resultado['Data'].isna().all():
            df_tempo = df_resultado.groupby('Data').agg({
                'Diferença Qtd': 'sum',
                'Documento': 'count'
            }).reset_index()
            df_tempo.columns = ['Data', 'Diferença Total', 'Número de Documentos']
            
            fig_linha = go.Figure()
            fig_linha.add_trace(go.Scatter(
                x=df_tempo['Data'],
                y=df_tempo['Diferença Total'],
                mode='lines+markers',
                name='Diferença Total',
                line=dict(color='#EF553B', width=2)
            ))
            fig_linha.update_layout(
                title='Evolução das Divergências ao Longo do Tempo',
                xaxis_title='Data',
                yaxis_title='Diferença Total de Quantidade',
                hovermode='x unified'
            )
            st.plotly_chart(fig_linha, use_container_width=True)
        
        st.markdown("---")
        
        # Tabela de resultados
        st.header("📋 Detalhamento dos Registros")
        
        # Adicionar coluna de status
        def definir_status(row):
            if row['Diferença Qtd'] == 0:
                return '✅ Conforme'
            elif row['Qtd Entrada'] == 0:
                return '❌ Sem Recebimento'
            else:
                return '⚠️ Divergência'
        
        df_resultado['Status'] = df_resultado.apply(definir_status, axis=1)
        
        # Reordenar colunas para incluir status
        colunas_exibicao = ['Data', 'Status', 'Produto (Saída)', 'Unidade Origem', 'Unidade Destino',
                           'Qtd Saída', 'Qtd Entrada', 'Diferença Qtd', 'Documento']
        df_exibicao = df_resultado[colunas_exibicao].copy()
        
        # Formatar data
        df_exibicao['Data'] = df_exibicao['Data'].dt.strftime('%d/%m/%Y')
        
        # Aplicar estilo condicional
        def highlight_divergencias(row):
            if row['Status'] == '❌ Sem Recebimento':
                return ['background-color: #ffcccc'] * len(row)
            elif row['Status'] == '⚠️ Divergência':
                return ['background-color: #fff4cc'] * len(row)
            else:
                return ['background-color: #ccffcc'] * len(row)
        
        st.dataframe(
            df_exibicao.style.apply(highlight_divergencias, axis=1),
            use_container_width=True,
            height=400
        )
        
        # Botão para download
        st.markdown("---")
        st.header("💾 Exportar Resultados")
        
        csv = df_resultado.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📥 Baixar Relatório em CSV",
            data=csv,
            file_name=f"relatorio_quantidades_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
        
        # Estatísticas adicionais
        with st.expander("📊 Estatísticas Detalhadas"):
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Por Unidade Origem")
                stats_origem = df_resultado.groupby('Unidade Origem').agg({
                    'Documento': 'count',
                    'Diferença Qtd': 'sum'
                }).reset_index()
                stats_origem.columns = ['Unidade', 'Total Documentos', 'Diferença Total']
                st.dataframe(stats_origem, use_container_width=True)
            
            with col2:
                st.subheader("Por Unidade Destino")
                stats_destino = df_resultado.groupby('Unidade Destino').agg({
                    'Documento': 'count',
                    'Diferença Qtd': 'sum'
                }).reset_index()
                stats_destino.columns = ['Unidade', 'Total Documentos', 'Diferença Total']
                st.dataframe(stats_destino, use_container_width=True)

else:
    st.info("👆 Por favor, faça o upload dos dois arquivos Excel na barra lateral para iniciar a análise.")
    
    st.markdown("""
    ### 📖 Como usar este dashboard:
    
    1. **Upload dos Arquivos**: Faça o upload dos arquivos Excel na barra lateral:
       - Empréstimos Concedidos (Saída)
       - Empréstimos Recebidos (Entrada)
    
    2. **Filtros**: Use os filtros na barra lateral para refinar sua análise:
       - Período de datas
       - Unidade de origem
       - Unidade de destino
       - Tipo de divergência
    
    3. **Análise**: Visualize:
       - Métricas gerais de divergências
       - Gráficos de distribuição e evolução
       - Tabela detalhada com todos os registros
       - Estatísticas por unidade
    
    4. **Export**: Baixe o relatório completo em formato CSV
    
    ### 🎯 Tipos de Status:
    - ✅ **Conforme**: Quantidade de saída = Quantidade de entrada
    - ⚠️ **Divergência**: Quantidade de saída ≠ Quantidade de entrada (mas houve recebimento)
    - ❌ **Sem Recebimento**: Quantidade de entrada = 0 (empréstimo não foi recebido)
    """)

# Rodapé
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray;'>Dashboard de Análise de Quantidades - Empréstimos Hospitalares</div>",
    unsafe_allow_html=True
)