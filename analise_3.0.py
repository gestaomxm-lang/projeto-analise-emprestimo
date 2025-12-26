import streamlit as st
import pandas as pd
import io
import re
import json
from difflib import SequenceMatcher
from datetime import datetime, timedelta
from pathlib import Path
import altair as alt
import plotly.express as px
import plotly.graph_objects as go
import pickle
import os
import analise_core  # Biblioteca de lógica central
import schedule
import time
import threading
import auto_analise

# --- Agendador em Background (Cron Job Simulado) ---
def run_pending_jobs():
    """Função rodada pela thread em background."""
    print("🕒 Iniciando loop do agendador em background...")
    while True:
        schedule.run_pending()
        time.sleep(60)

def job_atualizacao():
    """Tarefa que roda a cada hora."""
    print(f"⏰ [Auto-Update] Iniciando atualização agendada: {datetime.now()}")
    try:
        # Usa um container vazio pois não estamos no contexto da UI principal aqui
        # Apenas roda o fluxo backend
        sys.stdout = sys.__stdout__ # Garante log no console do servidor
        auto_analise.executar_fluxo_diario(baixar_email=True)
        print("✅ [Auto-Update] Concluído com sucesso.")
    except Exception as e:
        print(f"❌ [Auto-Update] Erro: {e}")

@st.cache_resource
def start_background_scheduler():
    """Inicia o agendador apenas uma vez (Singleton)."""
    # Agenda para rodar a cada 1 hora
    schedule.every(1).hours.do(job_atualizacao)
    # Também roda uma vez logo no início para garantir (opcional, já temos o run-on-load)
    
    # Inicia Thread
    t = threading.Thread(target=run_pending_jobs, daemon=True)
    t.start()
    return t

# Inicia o agendador
start_background_scheduler()


# Configuração da página
st.set_page_config(
    page_title="Análise de Empréstimos Hospitalares",
    page_icon="page_icon.png",
    layout="wide"
)

# --- Estilização Personalizada ---
st.markdown("""
    <style>
    .stAppHeader {
        background-color: #FFFFFF;
    }
    .block-container {
        padding-top: 5rem;
    }
    h1, h2, h3 {
        color: #001A72;
    }
    .stButton button {
        background-color: #E87722;
        color: white;
    }
    /* Cards de KPI */
    div[data-testid="stMetric"] {
        background-color: #F0F2F6;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #E87722;
    }
    /* Estilo do File Uploader */
    [data-testid="stFileUploader"] label {
        display: none !important;
    }
    
    /* Caixa visível em volta do botão de anexar arquivos */
    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed #E87722 !important;
        background-color: rgba(232, 119, 34, 0.06) !important;
        padding: 16px !important;
        min-height: 0px !important;
        border-radius: 10px !important;
    }
    
    /* Esconde especificamente os textos "Drag and drop" e "Limit" */
    [data-testid="stFileUploadDropzone"] > div {
        display: none !important;
    }
    
    /* Mantém apenas o botão visível */
    [data-testid="stFileUploadDropzone"] > section {
        display: block !important;
    }

    /* Estiliza o botão */
    [data-testid="stFileUploader"] button {
        font-size: 16px !important;
        background-color: #E87722 !important;
        color: transparent !important;
        border: none !important;
        padding: 10px 20px !important;
        border-radius: 5px !important;
        position: relative !important;
        width: 100% !important;
        display: block !important;
        visibility: visible !important;
    }
    [data-testid="stFileUploader"] button:hover {
        background-color: #d16615 !important;
        color: transparent !important;
    }
    [data-testid="stFileUploader"] button::after {
        content: "Anexar Arquivos" !important;
        color: white !important;
        position: absolute !important;
        left: 50% !important;
        top: 50% !important;
        transform: translate(-50%, -50%) !important;
        font-weight: bold !important;
        font-size: 16px !important;
    }
    </style>
""", unsafe_allow_html=True)

# JavaScript removido para evitar travamento do navegador (loop infinito no MutationObserver)
# O estilo CSS já deve ser suficiente para esconder os elementos


# --- Funções de Lógica de Negócio ---

# --- Regras de Negócio importadas de analise_core.py ---
# (Funções movidas para biblioteca externa para permitir automação)


def gerar_excel_bytes(df_in):
    # Cria cópia para não alterar o original da sessão
    df = df_in.copy()
    
    # Remove coluna auxiliar de data objeto se existir
    if 'Data_Obj' in df.columns:
        df = df.drop(columns=['Data_Obj'])
        
    # Remove timezones de colunas datetime (Excel não suporta)
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            try:
                df[col] = df[col].dt.tz_localize(None)
            except:
                pass
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Análise Completa")
        df[df['Status'].str.contains('Não Conforme', na=False)].to_excel(writer, index=False, sheet_name="Não Conformes")
        df[df['Status'].str.contains('Conforme', na=False) & ~df['Status'].str.contains('Não', na=False)].to_excel(writer, index=False, sheet_name="Conformes")
    return output.getvalue()

# --- Funções de Histórico ---

def get_history_dir():
    """Retorna o diretório de histórico, criando se não existir."""
    history_dir = Path("historico_analises")
    history_dir.mkdir(exist_ok=True)
    return history_dir

def save_analysis_to_history(df_resultado, stats, file_saida_name, file_entrada_name):
    """Salva uma análise no histórico."""
    try:
        history_dir = get_history_dir()
        timestamp = datetime.now()
        analysis_id = timestamp.strftime("%Y%m%d_%H%M%S")
        
        # Salva o DataFrame como CSV (mais leve que Excel)
        csv_path = history_dir / f"{analysis_id}_dados.csv"
        df_resultado.to_csv(csv_path, index=False, encoding='utf-8-sig')
        
        # Salva metadados
        metadata = {
            'id': analysis_id,
            'timestamp': timestamp.isoformat(),
            'data_formatada': timestamp.strftime("%d/%m/%Y %H:%M:%S"),
            'arquivo_saida': file_saida_name,
            'arquivo_entrada': file_entrada_name,
            'total_itens': len(df_resultado),
            'stats': stats
        }
        
        metadata_path = history_dir / f"{analysis_id}_metadata.json"
        with open(metadata_path, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, ensure_ascii=False, indent=2)
        
        return analysis_id
    except Exception as e:
        st.error(f"Erro ao salvar histórico: {e}")
        return None

def load_history_list():
    """Carrega lista de análises do histórico."""
    try:
        history_dir = get_history_dir()
        metadata_files = sorted(history_dir.glob("*_metadata.json"), reverse=True)
        
        history_list = []
        for metadata_file in metadata_files:
            with open(metadata_file, 'r', encoding='utf-8') as f:
                metadata = json.load(f)
                history_list.append(metadata)
        
        return history_list
    except Exception as e:
        st.error(f"Erro ao carregar histórico: {e}")
        return []

def load_analysis_from_history(analysis_id):
    """Carrega uma análise específica do histórico."""
    try:
        history_dir = get_history_dir()
        csv_path = history_dir / f"{analysis_id}_dados.csv"
        metadata_path = history_dir / f"{analysis_id}_metadata.json"
        
        if csv_path.exists() and metadata_path.exists():
            df = pd.read_csv(csv_path, encoding='utf-8-sig')
            
            # Garante que Data_Obj exista (para filtros) e Data esteja formatada
            if 'Data' in df.columns:
                # Converte para datetime primeiro para garantir normalização
                dt_series = pd.to_datetime(df['Data'], errors='coerce')
                
                # Se Data_Obj não existir (histórico antigo), cria
                if 'Data_Obj' not in df.columns:
                    df['Data_Obj'] = dt_series
                else:
                    df['Data_Obj'] = pd.to_datetime(df['Data_Obj'], errors='coerce')
                
                # Formata Data para string YYYY-MM-DD HH:MM:SS
                def _format_dt_str(x):
                    if pd.isna(x): return ""
                    try:
                        return x.strftime('%Y-%m-%d %H:%M:%S')
                    except:
                        return str(x)
                        
                df['Data'] = dt_series.apply(_format_dt_str)

            if 'Data Entrada' in df.columns:
                 df['Data Entrada'] = pd.to_datetime(df['Data Entrada'], errors='coerce').apply(lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notna(x) else x)
            
            with open(metadata_path, 'r', encoding='utf-8') as f:
                metadata = json.load(f)
            
            return df, metadata
        else:
            return None, None
    except Exception as e:
        st.error(f"Erro ao carregar análise: {e}")
        return None, None

def delete_analysis_from_history(analysis_id):
    """Remove uma análise do histórico."""
    try:
        history_dir = get_history_dir()
        csv_path = history_dir / f"{analysis_id}_dados.csv"
        metadata_path = history_dir / f"{analysis_id}_metadata.json"
        
        if csv_path.exists():
            csv_path.unlink()
        if metadata_path.exists():
            metadata_path.unlink()
        
        return True
    except Exception as e:
        st.error(f"Erro ao deletar análise: {e}")
        return False

# --- Interface Streamlit ---

# Inicializa chave do uploader se não existir
if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0


if 'df_resultado' not in st.session_state:
    st.session_state.df_resultado = None
if 'current_metadata' not in st.session_state:
    st.session_state.current_metadata = None

# --- Carregamento Automático da Análise Diária ---
import sys
from io import StringIO

class StdoutCapturer:
    def __init__(self, container):
        self.container = container
        self.buffer = StringIO()
        self.original_stdout = sys.stdout
        
    def __enter__(self):
        sys.stdout = self
        return self
        
    def __exit__(self, exc_type, exc_value, traceback):
        sys.stdout = self.original_stdout
        
    def write(self, s):
        self.original_stdout.write(s) # Maintain console log
        self.buffer.write(s)
        self.container.code(self.buffer.getvalue(), language="text")
        
    def flush(self):
        self.original_stdout.flush()

if st.session_state.df_resultado is None:
    try:
        daily_pkl = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados", "resultado_diario.pkl")
        if os.path.exists(daily_pkl):
            try:
                with open(daily_pkl, "rb") as f:
                    dados_auto = pickle.load(f)
                
                # Verifica data do arquivo
                data_proc = dados_auto['metadata']['data_processamento']
                is_today = data_proc.date() == datetime.now().date()
                
                if not is_today:
                     raise Exception("Dados desatualizados (não são de hoje)")
                     
                st.session_state.df_resultado = dados_auto['df']
                st.session_state.current_metadata = {
                    'arquivo_saida': dados_auto['metadata']['arquivo_saida'],
                    'arquivo_entrada': dados_auto['metadata']['arquivo_entrada'],
                    'data_formatada': data_proc.strftime("%d/%m/%Y %H:%M:%S"),
                    'modo': 'Automático 🤖'
                }
            except Exception:
                raise # Força cair no except abaixo que roda a automação
        else:
            raise FileNotFoundError("Arquivo de dados não existe") # Força cair no except

    except Exception as e:
        # Se der erro ou não existir, tenta RODAR O FLUXO AGORA (Self-Healing / Cloud Mode)
        st.warning(f"Dados locais não encontrados. Iniciando automação em nuvem...")
        
        log_container = st.empty()
        
        try:
            import auto_analise
            with StdoutCapturer(log_container):
                with st.spinner("Executando robô de análise..."):
                    sucesso = auto_analise.executar_fluxo_diario(baixar_email=True)
            
            if sucesso:
                # Tenta carregar novamente
                if os.path.exists(daily_pkl):
                    with open(daily_pkl, "rb") as f:
                        dados_auto = pickle.load(f)
                    
                    st.session_state.df_resultado = dados_auto['df']
                    st.session_state.current_metadata = {
                        'arquivo_saida': dados_auto['metadata']['arquivo_saida'],
                        'arquivo_entrada': dados_auto['metadata']['arquivo_entrada'],
                        'data_formatada': dados_auto['metadata']['data_processamento'].strftime("%d/%m/%Y %H:%M:%S"),
                        'modo': 'Automático (Sob Demanda) 🤖'
                    }
                    st.rerun() # Recarrega a página com os dados novos
                else:
                    st.error("Automação rodou mas arquivo não foi criado.")
            else:
                st.error("Falha na execução da automação. Verifique logs.")
        except Exception as e2:
             st.error(f"Erro crítico ao tentar rodar automação: {e2}")



col_logo, col_title, col_opts = st.columns([1, 4, 1])

with col_opts:
    st.empty() # Espaço vazio onde estava o botão


with col_logo:
    try:
        st.image("logo.png", use_container_width=True)
    except:
        st.warning("Logo?")

with col_title:
    st.title("Análise de Tranferências - Via Empréstimo")



# --- Status da Automação ---
if st.session_state.df_resultado is None:
    st.info("🤖 **Aguardando dados da automação...**")
    st.markdown("""
        O sistema processa novos arquivos automaticamente diariamente às 07:00.
        Se os dados não aparecerem, verifique se o serviço de agendamento está rodando.
        
        *Nenhuma ação manual é necessária.*
    """)


# --- Dashboard ---
if st.session_state.df_resultado is not None:
    df = st.session_state.df_resultado
    
    # Mostra informações da análise atual
    # Mostra informações da análise atual
    # Mostra informações do período apurado
    # Usa Data_Obj se existir (pois 'Data' virou string)
    col_data_ref = df['Data_Obj'] if 'Data_Obj' in df.columns else (pd.to_datetime(df['Data'], errors='coerce') if 'Data' in df.columns else None)
    
    min_date_apurado = col_data_ref.min() if col_data_ref is not None else None
    max_date_apurado = col_data_ref.max() if col_data_ref is not None else None
    
    if pd.notna(min_date_apurado) and pd.notna(max_date_apurado):
        periodo_str = f"{min_date_apurado.strftime('%d/%m/%Y')} até {max_date_apurado.strftime('%d/%m/%Y')}"
    else:
        periodo_str = "-"
        
    st.info(f"📅 **Período Apurado:** {periodo_str}")
    
    # --- Filtros no Topo dos Resultados ---
    with st.expander("🔍 Filtros do Dashboard", expanded=False):
        c_filt1, c_filt2, c_filt3 = st.columns(3)
        with c_filt1:
            # Filtro de Status
            status_options = df['Status'].unique()
            status_filter = st.multiselect("Status", status_options, default=status_options)
        
        with c_filt2:
            # Filtro de Unidade
            unidades = sorted(list(set(df['Unidade Origem'].unique()) | set(df['Unidade Destino'].unique())))
            unidade_filter = st.multiselect("Unidade (Origem/Destino)", unidades)
            
        with c_filt3:
            # Filtro de Data
            if col_data_ref is not None:
                min_date = col_data_ref.min()
                max_date = col_data_ref.max()
                date_range = st.date_input("Período", [min_date, max_date])
            else:
                date_range = []
    
    # Aplica Filtros
    df_filtered = df[df['Status'].isin(status_filter)]
    
    if unidade_filter:
        df_filtered = df_filtered[
            df_filtered['Unidade Origem'].isin(unidade_filter) | 
            df_filtered['Unidade Destino'].isin(unidade_filter)
        ]
        
    if len(date_range) == 2 and col_data_ref is not None:
        # Recupera índices que estão no range
        mask_data = (col_data_ref.dt.date >= date_range[0]) & (col_data_ref.dt.date <= date_range[1])
        df_filtered = df_filtered[mask_data]
    
    # --- Custom CSS & Helper for KPIs (Shared) ---
    st.markdown("""
    <style>
    .kpi-card {
        background-color: white;
        border-radius: 8px;
        padding: 15px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
        border-left: 5px solid #4A90E2;
        margin-bottom: 15px;
        min-height: 110px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        transition: transform 0.2s;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    .kpi-title {
        color: #666;
        font-size: 0.9em;
        font-weight: 600;
        margin-bottom: 5px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        white-space: normal;
        line-height: 1.2;
    }
    .kpi-value {
        color: #2C3E50;
        font-size: 1.5em;
        font-weight: 700;
        margin-bottom: 5px;
    }
    .kpi-subtitle {
        color: #888;
        font-size: 0.8em;
        line-height: 1.2;
    }
    .kpi-trend {
        font-weight: 600;
        padding: 1px 4px;
        border-radius: 3px;
        font-size: 0.85em;
    }
    .trend-up { color: #27ae60; background-color: #e8f8f0; }
    .trend-down { color: #e74c3c; background-color: #fdedec; }
    
    .b-blue { border-left-color: #3498db; }
    .b-green { border-left-color: #2ecc71; }
    .b-red { border-left-color: #e74c3c; }
    .b-orange { border-left-color: #f39c12; }
    .b-purple { border-left-color: #9b59b6; }
    </style>
    """, unsafe_allow_html=True)
    
    def kpi_card(title, value, subtitle=None, icon=None, color="b-blue", help_text=None):
        icon_html = f'<span style="margin-right:4px; font-size:1.1em">{icon}</span>' if icon else ''
        subtitle_html = f'<div class="kpi-subtitle">{subtitle}</div>' if subtitle else ''
        tooltip_attr = f'title="{help_text}"' if help_text else ''
        return f"""
        <div class="kpi-card {color}" {tooltip_attr}>
            <div class="kpi-title">{icon_html}{title}</div>
            <div class="kpi-value">{value}</div>
            {subtitle_html}
        </div>
        """

    def fmt_moeda(v):
        return f"R$ {v:_.2f}".replace('.', ',').replace('_', '.')
        
    def fmt_perc(v, total):
        return (v / total * 100) if total > 0 else 0

    # --- Balanço Financeiro do Período ---
    st.markdown("### Balanço Financeiro do Período")
    
    total_saida_periodo = df_filtered['Valor Saída (R$)'].sum()
    total_entrada_periodo = df_filtered[df_filtered['Valor Saída (R$)'].notna()]['Valor Entrada (R$)'].sum()
    valor_pendente = df_filtered[df_filtered['Status'].str.contains('Não Recebido')]['Valor Saída (R$)'].sum()
    valor_divergente_nc = df_filtered[df_filtered['Status'].str.contains('Não Conforme')]['Diferença (R$)'].abs().sum()
    
    col_balanco1, col_balanco2, col_balanco3, col_balanco4 = st.columns(4)
    
    with col_balanco1:
        st.markdown(kpi_card("Total Saída", fmt_moeda(total_saida_periodo), "Enviado", "📤", "b-blue", help_text="Soma do valor de todos os itens de saída no período"), unsafe_allow_html=True)
        
    with col_balanco2:
        st.markdown(kpi_card("Total Entrada", fmt_moeda(total_entrada_periodo), "Recebido", "📥", "b-green", help_text="Soma do valor de todos os itens de entrada no período"), unsafe_allow_html=True)
        
    with col_balanco3:
        p_pend = fmt_perc(valor_pendente, total_saida_periodo)
        st.markdown(kpi_card("Pendentes", fmt_moeda(valor_pendente), f"<span class='kpi-trend trend-down'>⬇ {p_pend:.1f}% do total</span>", "⏳", "b-orange", help_text="Valor total dos itens que ainda não foram recebidos"), unsafe_allow_html=True)

    with col_balanco4:
        p_div = fmt_perc(valor_divergente_nc, total_entrada_periodo)
        st.markdown(kpi_card("Divergência Itens Recebidos", fmt_moeda(valor_divergente_nc), f"<span class='kpi-trend trend-down'>⚠️ {p_div:.1f}% da entrada</span>", "≠", "b-red", help_text="Soma absoluta das diferenças dos itens recebidos com divergência"), unsafe_allow_html=True)
    
    st.divider()

    # --- KPIs Operacionais ---
    st.markdown("### Indicadores Operacionais")

    # --- Cálculos Operacionais ---
    
    total_analisado = len(df_filtered)
    
    total_conforme = len(df_filtered[df_filtered['Status'].str.contains('Conforme') & ~df_filtered['Status'].str.contains('Não')])
    total_nao_conforme = len(df_filtered[df_filtered['Status'].str.contains('Não Conforme')])
    total_pendente = len(df_filtered[df_filtered['Status'].str.contains('Não Recebido')])
    
    # Conta itens com divergência de quantidade (excluindo não encontrados/nulos)
    qtd_divergente_count = len(df_filtered[
        (df_filtered['Diferença Qtd'].notna()) & 
        (df_filtered['Diferença Qtd'] != 0)
    ])
    
    # Conta itens com entrada anterior à saída (Tempo Recebimento negativo)
    entradas_anteriores_count = len(df_filtered[
        (df_filtered['Tempo Recebimento (Horas)'].notna()) & 
        (df_filtered['Tempo Recebimento (Horas)'] < 0)
    ])
    
    # Cálculo Média Tempo Recebimento
    if 'Tempo Recebimento (Horas)' in df_filtered.columns:
        df_tempo_calc = df_filtered[
            ~df_filtered['Status'].str.contains('Não Recebido') & 
            df_filtered['Tempo Recebimento (Horas)'].notna() &
            (df_filtered['Tempo Recebimento (Horas)'] >= 0)
        ]
        media_tempo = df_tempo_calc['Tempo Recebimento (Horas)'].mean()
        media_tempo_str = f"{media_tempo:.1f} horas" if pd.notna(media_tempo) else "-"
        
        # Formata visualmente para HH:MM:SS
        def _format_hours_hms(h):
            if pd.isna(h): return "-"
            try:
                total_seconds = int(h * 3600)
                sign = "-" if total_seconds < 0 else ""
                total_seconds = abs(total_seconds)
                hours, remainder = divmod(total_seconds, 3600)
                minutes, seconds = divmod(remainder, 60)
                return f"{sign}{hours:02d}:{minutes:02d}:{seconds:02d}"
            except:
                return h
        df_filtered['Tempo Recebimento (Horas)'] = df_filtered['Tempo Recebimento (Horas)'].apply(_format_hours_hms)
    else:
        media_tempo_str = "-"

    # Layout Operacional - Linha 1 (4 itens)
    op1, op2, op3, op4 = st.columns(4)
    with op1: 
        st.markdown(kpi_card("Total Itens", f"{total_analisado:_.0f}".replace("_", "."), "Processados", "📦", "b-blue"), unsafe_allow_html=True)
    with op2: 
        p_c = (total_conforme / total_analisado * 100) if total_analisado > 0 else 0
        st.markdown(kpi_card("Conformes", f"{total_conforme:_.0f}".replace("_", "."), f"✓ {p_c:.1f}%", "✅", "b-green"), unsafe_allow_html=True)
    with op3: 
        p_nc = (total_nao_conforme / total_analisado * 100) if total_analisado > 0 else 0
        st.markdown(kpi_card("Não Conformes", f"{total_nao_conforme:_.0f}".replace("_", "."), f"! {p_nc:.1f}%", "❌", "b-red"), unsafe_allow_html=True)
    with op4: 
        st.markdown(kpi_card("Itens Pendentes", f"{total_pendente:_.0f}".replace("_", "."), "Sem Entrada", "⏳", "b-orange"), unsafe_allow_html=True)

    # Layout Operacional - Linha 2 Centrada (3 itens)
    # [vazio 0.5] [KPI] [KPI] [KPI] [vazio 0.5]
    _, op5, op6, op7, _ = st.columns([0.5, 1, 1, 1, 0.5])
    
    with op5: st.markdown(kpi_card("Entradas Inferiores a Saída (Data Recebimento) ", f"{entradas_anteriores_count:_.0f}".replace("_", "."), "Entrada < Saída", "📅", "b-red"), unsafe_allow_html=True)
    with op6: st.markdown(kpi_card("Divergência de Quantidade", f"{qtd_divergente_count:_.0f}".replace("_", "."), "Entrada ≠ Saída", "🔢", "b-red"), unsafe_allow_html=True)
    with op7: st.markdown(kpi_card("Tempo Médio Recebimento", media_tempo_str, "Ciclo Validado", "⏱️", "b-purple"), unsafe_allow_html=True)
    
    st.divider()
    
    # Expander com detalhamento de divergências de quantidade
    if qtd_divergente_count > 0:
        with st.expander(f"📋 Detalhamento de Divergências de Quantidade ({qtd_divergente_count} itens)", expanded=False):
            df_qtd_div = df_filtered[
                (df_filtered['Diferença Qtd'].notna()) & 
                (df_filtered['Diferença Qtd'] != 0)
            ].copy()
            
            # Seleciona e renomeia colunas relevantes
            df_qtd_div_display = df_qtd_div[[
                'Data', 'Produto (Saída)', 'Unidade Origem', 'Unidade Destino',
                'Qtd Saída', 'Qtd Entrada', 'Diferença Qtd', 'Documento'
            ]].copy()
            
            # Formata a diferença com sinal
            df_qtd_div_display['Diferença Qtd'] = df_qtd_div_display['Diferença Qtd'].apply(
                lambda x: f"{x:+.0f}" if pd.notna(x) else "-"
            )
            
            st.dataframe(
                df_qtd_div_display,
                use_container_width=True,
                hide_index=True,
                height=400
            )
            
            # Resumo
            total_falta = df_qtd_div[df_qtd_div['Diferença Qtd'] < 0]['Diferença Qtd'].sum()
            total_sobra = df_qtd_div[df_qtd_div['Diferença Qtd'] > 0]['Diferença Qtd'].sum()
            
            col_res1, col_res2, col_res3 = st.columns(3)
            with col_res1:
                st.metric("Itens com Falta", f"{len(df_qtd_div[df_qtd_div['Diferença Qtd'] < 0])}", 
                         delta=f"{total_falta:.0f} unidades", delta_color="inverse")
            with col_res2:
                st.metric("Itens com Sobra", f"{len(df_qtd_div[df_qtd_div['Diferença Qtd'] > 0])}", 
                         delta=f"+{total_sobra:.0f} unidades", delta_color="off")
            with col_res3:
                st.metric("Divergência Total", f"{abs(total_falta) + total_sobra:.0f} unidades")
                
    st.divider()
    
    # --- Gráfico Top 5 Hospitais com Pendências ---
    st.markdown("### Top 5 Hospitais com Pendências de Entrada (Envios não Recebidos)")
    
    # Filtra itens pendentes (Não Recebido)
    df_pendentes = df_filtered[df_filtered['Status'].str.contains('Não Recebido', na=False)]
    
    if not df_pendentes.empty:
        # Agrupa por hospital (Unidade Destino do envio) e conta
        top5_pendentes = df_pendentes['Unidade Destino'].value_counts().head(5).reset_index()
        top5_pendentes.columns = ['Hospital', 'Quantidade']
        
        # Cria gráfico horizontal com Altair para customização
        base = alt.Chart(top5_pendentes).encode(
            x=alt.X('Quantidade', title='Quantidade de Pendências'),
            y=alt.Y('Hospital', sort='-x', title=None, axis=alt.Axis(labelLimit=400)), # Aumenta limite para nome não cortar
            tooltip=['Hospital', 'Quantidade']
        )
        
        bars = base.mark_bar(color='#E87722').encode() # Laranja tom mais profissional/premium
        
        text = base.mark_text(
            align='left',
            baseline='middle',
            dx=3  # Nudges text to right of bar
        ).encode(
            text='Quantidade'
        )
        
        st.altair_chart((bars + text).properties(height=300), use_container_width=True)
    else:
        st.info("Não há pendências de entrada registradas no período selecionado.")
    
    # Define cor do texto dos gráficos (sempre claro)
    chart_text_color = '#001A72'
    chart_grid_color = 'rgba(128,128,128,0.2)'

    # --- Gráficos Premium ---
    col_chart1, col_chart2 = st.columns(2)
    
    with col_chart1:
        st.markdown("#### Status de Recebimento")
        
        # Conta status
        status_counts = df_filtered['Status'].value_counts()
        
        # Remove emojis das labels e define cores
        clean_labels = [label.replace('✅ ', '').replace('❌ ', '').replace('⚠️ ', '') for label in status_counts.index]
        
        # Mapeamento de cores fixo
        color_map = {
            'Conforme': '#00C853',      # Verde
            'Não Conforme': '#FF4444',  # Vermelho
            'Não Recebido': '#FF9800'   # Laranja
        }
        
        # Gera lista de cores na ordem dos dados
        chart_colors = [color_map.get(label, '#999999') for label in clean_labels]
        
        # Gráfico de rosca com Plotly - texto otimizado
        fig_status = go.Figure(data=[go.Pie(
            labels=clean_labels,
            values=status_counts.values,
            hole=0.6,
            marker=dict(
                colors=chart_colors,
                line=dict(color='white', width=3)
            ),
            textposition='outside',
            textinfo='percent', # Apenas percentual no gráfico
            textfont=dict(size=12, family="Arial", color=chart_text_color),
            insidetextorientation='radial',
            pull=[0.05] * len(status_counts),  # Separa levemente todas as fatias
            hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>'
        )])
        
        fig_status.update_layout(
            showlegend=True,  # Exibe legenda com ícones (bolinhas/quadrados)
            legend=dict(
                orientation="v",
                yanchor="bottom",
                y=0,
                xanchor="right",
                x=1,
                font=dict(size=11, color=chart_text_color),
                bgcolor="rgba(0,0,0,0)" # Fundo transparente
            ),
            height=380,
            margin=dict(t=20, b=20, l=40, r=40), # Margens maiores para não cortar texto
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Arial", size=12, color=chart_text_color)
        )
        
        st.plotly_chart(fig_status, use_container_width=True)
        
    with col_chart2:
        st.markdown("#### Top 5 Hospitais com Divergências nas Quantidades Recebidas")
        
        df_div = df_filtered[df_filtered['Status'].str.contains('Não Conforme')]
        if not df_div.empty:
            # Combina origem e destino
            hospitais_div = pd.concat([
                df_div['Unidade Origem'].value_counts(),
                df_div['Unidade Destino'].value_counts()
            ]).groupby(level=0).sum().sort_values(ascending=False).head(5).reset_index()
            hospitais_div.columns = ['Hospital', 'Quantidade']
            
            # Cria gráfico Altair (consistente com o gráfico de Pendências)
            base_div = alt.Chart(hospitais_div).encode(
                x=alt.X('Quantidade', title='Quantidade de Divergências'),
                y=alt.Y('Hospital', sort='-x', title=None, axis=alt.Axis(labelLimit=400)),
                tooltip=['Hospital', 'Quantidade']
            )
            
            bars_div = base_div.mark_bar(color='#E87722').encode()
            
            text_div = base_div.mark_text(
                align='left',
                baseline='middle',
                dx=3
            ).encode(
                text='Quantidade'
            )
            
            st.altair_chart((bars_div + text_div).properties(height=350), use_container_width=True)
        else:
            st.info("Nenhuma divergência por hospital!")
    
    st.divider()

    # --- Tabela Detalhada ---
    st.subheader("Detalhamento dos Dados")
    
    st.dataframe(
        df_filtered,
        use_container_width=True,
        column_config={
            "Data": st.column_config.TextColumn("Data"),
            "Valor Saída (R$)": st.column_config.NumberColumn("Valor Saída", format="R$ %.2f"),
            "Valor Entrada (R$)": st.column_config.NumberColumn("Valor Entrada", format="R$ %.2f"),
            "Diferença (R$)": st.column_config.NumberColumn("Diferença", format="R$ %.2f"),
            "Status": st.column_config.TextColumn("Status"),
        },
        hide_index=True
    )
    
    # Download
    excel_data = gerar_excel_bytes(df_filtered)
    st.download_button(
        label="Baixar Dados Filtrados (Excel)",
        data=excel_data,
        file_name="analise_dashboard.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

