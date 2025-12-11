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

# Configuração da página
st.set_page_config(
    page_title="Análise de Empréstimos Hospitalares",
    page_icon="🏥",
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

# JavaScript para remover textos indesejados
st.markdown("""
    <script>
    // Remove textos do file uploader
    function cleanFileUploader() {
        const dropzone = document.querySelector('[data-testid="stFileUploadDropzone"]');
        if (dropzone) {
            // Remove todos os spans e smalls (textos de instrução)
            const spans = dropzone.querySelectorAll('span');
            const smalls = dropzone.querySelectorAll('small');
            spans.forEach(el => {
                if (!el.closest('button')) {
                    el.style.display = 'none';
                }
            });
            smalls.forEach(el => el.style.display = 'none');
        }
    }
    
    // Executa quando a página carrega
    cleanFileUploader();
    
    // Executa novamente após um pequeno delay (para garantir que o Streamlit terminou de renderizar)
    setTimeout(cleanFileUploader, 100);
    setTimeout(cleanFileUploader, 500);
    setTimeout(cleanFileUploader, 1000);
    
    // Observa mudanças no DOM e reaplica
    const observer = new MutationObserver(cleanFileUploader);
    observer.observe(document.body, { childList: true, subtree: true });
    </script>
""", unsafe_allow_html=True)

# --- Funções de Lógica de Negócio ---

def extrair_numeros(documento):
    """Extrai a primeira sequência numérica encontrada no documento."""
    match = re.search(r'\d+', str(documento))
    return match.group() if match else ""

def _parse_date_column(series):
    """Converte a coluna de datas para datetime."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_datetime(series, unit='d', origin='1899-12-30', errors='coerce')
    return pd.to_datetime(series, dayfirst=True, errors='coerce')

def normalizar_valor_numerico(valor):
    """Normaliza valores numéricos removendo separadores de milhar (ponto) e convertendo para float.
    
    Exemplos:
        "120.000" -> 120.0
        "1.234.567" -> 1234567.0
        "120,50" -> 120.50
        "120" -> 120.0
        "120.5" -> 120.5 (decimal preservado)
    """
    if pd.isna(valor) or valor == '' or valor is None:
        return 0.0
    
    # Se já é numérico, retorna direto
    if isinstance(valor, (int, float)):
        return float(valor)
    
    # Converte para string
    valor_str = str(valor).strip()
    
    # Remove espaços
    valor_str = valor_str.replace(' ', '')
    
    # Se vazio após limpeza, retorna 0
    if not valor_str:
        return 0.0
    
    # Detecta formato brasileiro (vírgula como decimal)
    if ',' in valor_str and '.' in valor_str:
        # Formato: 1.234,56 -> remove pontos (milhar) e substitui vírgula por ponto
        valor_str = valor_str.replace('.', '').replace(',', '.')
    elif ',' in valor_str:
        # Formato: 1234,56 -> substitui vírgula por ponto
        valor_str = valor_str.replace(',', '.')
    elif '.' in valor_str:
        # Pode ser separador de milhar (120.000) ou decimal (120.5)
        partes = valor_str.split('.')
        if len(partes) > 2:
            # Múltiplos pontos = separador de milhar (ex: 1.234.567)
            valor_str = valor_str.replace('.', '')
        elif len(partes) == 2:
            parte_decimal = partes[1]
            # Se termina com exatamente 3 zeros (120.000) ou mais de 3 dígitos, é milhar
            if parte_decimal == '000' or len(parte_decimal) > 3:
                # Separador de milhar
                valor_str = valor_str.replace('.', '')
            # Caso contrário, mantém como decimal (120.5, 120.50, etc.)
    
    try:
        return float(valor_str)
    except (ValueError, TypeError):
        return 0.0

def normalizar_unidade_medida(texto):
    """Normaliza unidades de medida para formato padrão."""
    mapeamento = {
        'GR': 'G',
        'GRAMA': 'G',
        'GRAMAS': 'G',
        'MILIGRAMA': 'MG',
        'MILIGRAMAS': 'MG',
        'MILILITRO': 'ML',
        'MILILITROS': 'ML',
        'MICROGRAMA': 'MCG',
        'MICROGRAMAS': 'MCG',
        'UNIDADE': 'UI',
        'UNIDADES': 'UI',
        'LITRO': 'L',
        'LITROS': 'L',
        'METRO': 'M',
        'METROS': 'M',
        'CENTIMETRO': 'CM',
        'CENTIMETROS': 'CM',
        'MILIMETRO': 'MM',
        'MILIMETROS': 'MM'
    }
    
    texto_norm = texto.upper()
    for variacao, padrao in mapeamento.items():
        texto_norm = re.sub(r'\b' + variacao + r'\b', padrao, texto_norm)
    
    return texto_norm

def normalizar_dimensao(dimensao_str):
    """Normaliza dimensões para comparação consistente."""
    # Remove espaços
    dimensao_str = re.sub(r'\s+', '', dimensao_str.upper())
    
    # Extrai números
    numeros = re.findall(r'\d+\.?\d*', dimensao_str)
    
    # Normaliza removendo zeros desnecessários e ordena
    numeros_norm = [str(float(n)) for n in numeros if n]
    
    # Ordena para comparar 25X7 com 7X25
    return 'X'.join(sorted(numeros_norm))

def extrair_e_normalizar_concentracao(descricao):
    """Extrai concentração e normaliza para formato comparável."""
    descricao = normalizar_unidade_medida(descricao)
    
    # Padrões expandidos para capturar diferentes formatos
    padroes = [
        r'(\d+[,.]?\d*)\s*(MG|G|ML|MCG|UI|L|%)\s*/\s*(\d+[,.]?\d*)\s*(MG|G|ML|MCG|UI|L)',  # 50MG/5ML
        r'(\d+[,.]?\d*)\s*(MG|G|ML|MCG|UI|L|%)(?!\s*/)',  # 50MG
    ]
    
    concentracoes = []
    for padrao in padroes:
        matches = re.findall(padrao, descricao)
        for match in matches:
            if isinstance(match, tuple):
                # Junta os elementos da tupla
                conc_str = ''.join(str(c) for c in match).replace(',', '.')
                concentracoes.append(conc_str)
    
    return ' '.join(concentracoes) if concentracoes else ''

def extrair_componentes_produto(descricao):
    """Extrai componentes principais do produto para matching inteligente."""
    descricao = str(descricao).upper().strip()
    
    # Aplica normalização de unidades primeiro
    descricao = normalizar_unidade_medida(descricao)
    
    componentes = {
        'original': descricao,
        'normalizado': '',
        'principio_ativo': '',
        'concentracao': '',
        'apresentacao': '',
        'quantidade': '',
        'unidade_medida': '',
        'dimensao': '',
        'palavras_chave': []
    }
    
    # Remove pontuações e normaliza
    texto_limpo = re.sub(r'[^\w\s]', ' ', descricao)
    texto_limpo = re.sub(r'\s+', ' ', texto_limpo).strip()
    componentes['normalizado'] = texto_limpo
    
    # Extrai concentração usando função melhorada
    componentes['concentracao'] = extrair_e_normalizar_concentracao(descricao)
    
    # Extrai apresentação
    apresentacoes = [
        'AMPOLA', 'AMP', 'COMPRIMIDO', 'COMP', 'CP', 'CAPSULA', 'CAPS',
        'FRASCO', 'FR', 'SERINGA', 'SER', 'BOLSA', 'ENVELOPE', 'ENV',
        'TUBO', 'BISNAGA', 'SACHÊ', 'SACHE', 'BLISTER', 'CARTELA',
        'POTE', 'VIDRO', 'UNIDADE', 'UN', 'CAIXA', 'CX'
    ]
    for apres in apresentacoes:
        if apres in descricao:
            componentes['apresentacao'] = apres
            break
    
    # Extrai quantidade (ex: CX C/ 10, C/50, X10, etc)
    qtd_match = re.search(r'(?:C/|C |X|COM )\s*(\d+)', descricao)
    if qtd_match:
        componentes['quantidade'] = qtd_match.group(1)
        
    # Extrai dimensões (ex: 13X4,5, 25X7, 40X12, 0.70X25MM)
    dimensoes = re.search(r'\d+\.?\d*\s*[xX]\s*\d+\.?\d*', descricao)
    if dimensoes:
        componentes['dimensao'] = normalizar_dimensao(dimensoes.group())
    
    # Extrai palavras-chave (remove stopwords médicas comuns)
    stopwords = [
        'DE', 'DA', 'DO', 'COM', 'PARA', 'EM', 'A', 'O', 'E', 'C/',
        'SOLUCAO', 'SOL', 'INJETAVEL', 'INJ', 'ORAL', 'USO', 'ADULTO',
        'PEDIATRICO', 'ESTERIL', 'DESCARTAVEL', 'DESC'
    ]
    palavras = texto_limpo.split()
    palavras_chave = [p for p in palavras if p not in stopwords and len(p) > 2]
    componentes['palavras_chave'] = palavras_chave[:5]
    
    if len(palavras_chave) > 0:
        componentes['principio_ativo'] = ' '.join(palavras_chave[:2])
    
    return componentes

def calcular_similaridade_precalc(comp1, comp2, ignore_penalties=False):
    """Calcula similaridade usando componentes pré-calculados."""
    score = 0
    detalhes = []
    
    # 0. Verificação de Sinônimos (Bônus imediato)
    sinonimos = {
        # Equipamentos
        'AVENTAL': ['CAPOTE', 'AVENTAL', 'JALECO'],
        'CAPOTE': ['AVENTAL', 'CAPOTE', 'JALECO'],
        'JALECO': ['AVENTAL', 'CAPOTE', 'JALECO'],
        
        # Materiais
        'ALGODAO': ['POLYCOT', 'ALGODAO', 'COTTON'],
        'POLYCOT': ['ALGODAO', 'POLYCOT', 'COTTON'],
        'COTTON': ['ALGODAO', 'POLYCOT', 'COTTON'],
        'GAZE': ['COMPRESSA', 'GAZE'],
        'COMPRESSA': ['GAZE', 'COMPRESSA'],
        
        # Soluções
        'SORO': ['SOLUCAO', 'SORO', 'SOL'],
        'SOLUCAO': ['SORO', 'SOLUCAO', 'SOL'],
        'SOL': ['SORO', 'SOLUCAO', 'SOL'],
        'SALINA': ['NACL', 'CLORETO', 'SALINA', 'SF'],
        'NACL': ['SALINA', 'CLORETO', 'NACL', 'SF'],
        'SF': ['SALINA', 'NACL', 'CLORETO', 'SF'],
        
        # Formas farmacêuticas
        'AMPOLA': ['AMP', 'AMPOLA', 'FRAMP', 'FRASCOAMPOLA'],
        'AMP': ['AMPOLA', 'AMP', 'FRAMP', 'FRASCOAMPOLA'],
        'FRAMP': ['AMP', 'AMPOLA', 'FRAMP', 'FRASCOAMPOLA'],
        'FRASCOAMPOLA': ['AMP', 'AMPOLA', 'FRAMP', 'FRASCOAMPOLA'],
        'COMPRIMIDO': ['COMP', 'CP', 'COMPRIMIDO', 'DRAGEA'],
        'COMP': ['COMPRIMIDO', 'COMP', 'CP', 'DRAGEA'],
        'CP': ['COMPRIMIDO', 'COMP', 'CP', 'DRAGEA'],
        'CAPSULA': ['CAPS', 'CAPSULA', 'CAP'],
        'CAPS': ['CAPSULA', 'CAPS', 'CAP'],
        'CAP': ['CAPSULA', 'CAPS', 'CAP'],
        
        # Vias de administração
        'INJETAVEL': ['INJ', 'INJETAVEL', 'IV', 'IM', 'SC'],
        'INJ': ['INJETAVEL', 'INJ', 'IV', 'IM', 'SC'],
        'ORAL': ['VO', 'ORAL', 'BUCAL'],
        'VO': ['ORAL', 'VO', 'BUCAL'],
        
        # Princípios ativos comuns
        'DIPIRONA': ['METAMIZOL', 'DIPIRONA', 'NOVALGINA'],
        'METAMIZOL': ['DIPIRONA', 'METAMIZOL', 'NOVALGINA'],
        'PARACETAMOL': ['ACETAMINOFENO', 'PARACETAMOL'],
        'ACETAMINOFENO': ['PARACETAMOL', 'ACETAMINOFENO'],
        'OMEPRAZOL': ['OMEPRAZOL', 'LOSEC'],
        'DICLOFENACO': ['DICLOFENACO', 'VOLTAREN', 'CATAFLAM'],
        'GLICOSE': ['DEXTROSE', 'GLICOSE'],
        'DEXTROSE': ['GLICOSE', 'DEXTROSE'],
    }
    
    tem_sinonimo = False
    for termo, lista_sin in sinonimos.items():
        if termo in comp1['normalizado'] and any(s in comp2['normalizado'] for s in lista_sin):
            tem_sinonimo = True
            break
            
    if tem_sinonimo:
        score += 15
        detalhes.append("Sinônimo:✓")
    
    # 1. Similaridade textual geral (30%)
    sim_geral = SequenceMatcher(None, comp1['normalizado'], comp2['normalizado']).ratio()
    score += sim_geral * 30
    detalhes.append(f"Texto:{sim_geral:.0%}")
    
    # 2. Princípio ativo (35%)
    if comp1['principio_ativo'] and comp2['principio_ativo']:
        sim_principio = SequenceMatcher(None, comp1['principio_ativo'], comp2['principio_ativo']).ratio()
        score += sim_principio * 35
        detalhes.append(f"Princípio:{sim_principio:.0%}")
    
    # 3. Concentração (20%)
    if comp1['concentracao'] and comp2['concentracao']:
        c1 = comp1['concentracao'].replace(' ', '').upper()
        c2 = comp2['concentracao'].replace(' ', '').upper()
        
        # Comparação exata
        if c1 == c2:
            score += 20
            detalhes.append(f"Conc:✓")
        else:
            # Extrai números para comparação numérica
            nums1 = re.findall(r'\d+\.?\d*', c1)
            nums2 = re.findall(r'\d+\.?\d*', c2)
            
            # Se tem números em comum, pode ser variação do mesmo produto
            nums_comum = set(nums1) & set(nums2)
            
            if nums_comum and len(nums_comum) >= len(nums1) * 0.5:
                # Pelo menos 50% dos números batem
                score += 15
                detalhes.append(f"Conc:~")
            else:
                # Tenta similaridade textual como fallback
                sim_conc = SequenceMatcher(None, c1, c2).ratio()
                if sim_conc > 0.7:
                    score += sim_conc * 15
                    detalhes.append(f"Conc:~{sim_conc:.0%}")
                elif not ignore_penalties:
                    # Penaliza concentrações claramente diferentes
                    score -= 25
                    detalhes.append(f"Conc:Mismatch")
    
    # 4. Dimensão (bônus ou penalidade) - Importante para agulhas
    if 'dimensao' in comp1 and 'dimensao' in comp2 and comp1['dimensao'] and comp2['dimensao']:
        # Usa dimensões já normalizadas (ordenadas e sem zeros extras)
        d1_norm = comp1['dimensao']
        d2_norm = comp2['dimensao']
        
        if d1_norm == d2_norm:
            score += 15
            detalhes.append(f"Dim:✓")
        else:
            # Verifica se os números individuais batem
            nums1 = set(d1_norm.split('X'))
            nums2 = set(d2_norm.split('X'))
            comum = nums1 & nums2
            
            if len(comum) >= 2:  # Pelo menos 2 números batem
                score += 10
                detalhes.append(f"Dim:~")
            elif len(comum) >= 1:  # Pelo menos 1 número bate
                score += 5
                detalhes.append(f"Dim:part")
            elif not ignore_penalties:
                # Dimensões completamente diferentes são críticas
                score -= 15
                detalhes.append(f"Dim:Mismatch")
    
    # 5. Apresentação (10%)
    if comp1['apresentacao'] and comp2['apresentacao']:
        if comp1['apresentacao'] == comp2['apresentacao']:
            score += 10
            detalhes.append(f"Apres:✓")
        else:
            equiv_apresentacao = {
                'AMPOLA': ['AMP', 'AMPOLA', 'FR/AMP', 'FRASCO/AMPOLA'],
                'FR/AMP': ['AMP', 'AMPOLA', 'FR/AMP', 'FRASCO/AMPOLA'],
                'COMPRIMIDO': ['COMP', 'CP', 'COMPRIMIDO'],
                'CAPSULA': ['CAPS', 'CAPSULA'],
                'FRASCO': ['FR', 'FRASCO', 'FR/AMP'],
                'SERINGA': ['SER', 'SERINGA']
            }
            match_apres = False
            for grupo in equiv_apresentacao.values():
                if comp1['apresentacao'] in grupo and comp2['apresentacao'] in grupo:
                    score += 10
                    detalhes.append(f"Apres:equiv")
                    match_apres = True
                    break
            
            if not match_apres and not ignore_penalties:
                 # Apresentações diferentes (ex: Comprimido vs Ampola)
                 score -= 10
                 detalhes.append(f"Apres:Mismatch")

    # 6. Palavras-chave (5%)
    palavras_comum = set(comp1['palavras_chave']) & set(comp2['palavras_chave'])
    if palavras_comum:
        perc_comum = len(palavras_comum) / max(len(comp1['palavras_chave']), len(comp2['palavras_chave']))
        score += perc_comum * 5
        detalhes.append(f"Palavras:{len(palavras_comum)}")
    
    return score, ' | '.join(detalhes)

def validar_match_quantidade(qtd_saida, qtd_entrada, score_produto, doc_match):
    """Valida se o match de quantidade faz sentido para evitar falsos positivos.
    
    Retorna: (is_valid, diferenca_qtd)
    """
    # Se as quantidades são exatamente iguais (com tolerância float), sempre válido
    if abs(qtd_saida - qtd_entrada) < 0.01:
        return True, 0.0
    
    # Calcula diferença percentual
    qtd_max = max(qtd_saida, qtd_entrada)
    if qtd_max == 0:
        return True, 0.0
    
    perc_diff = abs(qtd_saida - qtd_entrada) / qtd_max * 100
    
    # Se tem documento correspondente e produto muito similar, aceita variação maior
    if doc_match and score_produto >= 85:
        # Tolera até 20% de diferença se doc e produto batem bem
        if perc_diff <= 20:
            return True, qtd_saida - qtd_entrada
    
    # Se não tem documento, exige quantidade mais próxima
    if not doc_match:
        # Só aceita se diferença for pequena (<10%)
        if perc_diff > 10:
            return False, None
    
    # Caso padrão: aceita o match
    return True, qtd_saida - qtd_entrada

def eh_casa_portugal(unidade):
    unidade_norm = str(unidade).upper().strip()
    return 'CASA' in unidade_norm and 'PORTUGAL' in unidade_norm

def analisar_itens(df_saida, df_entrada, limiar_similaridade=65, progress_bar=None):
    analise = []
    entradas_processadas = set()
    
    # Determina o período analisado com base nas saídas
    periodo_inicio = df_saida['data'].min() if 'data' in df_saida.columns else None
    periodo_fim = df_saida['data'].max() if 'data' in df_saida.columns else None
    
    # Normaliza colunas
    for df in [df_saida, df_entrada]:
        df['documento'] = df['documento'].astype(str).str.strip()
        df['ds_produto'] = df['ds_produto'].astype(str).str.strip()
        df['unidade_origem'] = df['unidade_origem'].astype(str).str.strip()
        df['unidade_destino'] = df['unidade_destino'].astype(str).str.strip()
    
    # --- PRÉ-CÁLCULO MASSIVO (OTIMIZAÇÃO) ---
    if progress_bar:
        progress_bar.progress(0.05, text="Pré-processando dados...")
    
    # Calcula componentes apenas uma vez para cada linha
    df_saida['comps'] = df_saida['ds_produto'].apply(extrair_componentes_produto)
    df_entrada['comps'] = df_entrada['ds_produto'].apply(extrair_componentes_produto)
    
    df_saida['doc_num'] = df_saida['documento'].apply(extrair_numeros)
    df_entrada['doc_num'] = df_entrada['documento'].apply(extrair_numeros)
    
    df_saida['destino_cp'] = df_saida['unidade_destino'].apply(eh_casa_portugal)
    
    # Pré-calcula strings normalizadas para comparação rápida
    df_saida['origem_norm'] = df_saida['unidade_origem'].str.upper().str.strip()
    df_saida['destino_norm'] = df_saida['unidade_destino'].str.upper().str.strip()
    df_entrada['origem_norm'] = df_entrada['unidade_origem'].str.upper().str.strip()
    df_entrada['destino_norm'] = df_entrada['unidade_destino'].str.upper().str.strip()
    
    # Cria índice por documento para busca O(1)
    doc_index = {}
    for idx, row in df_entrada.iterrows():
        doc = row['doc_num']
        if doc and doc != '':
            if doc not in doc_index:
                doc_index[doc] = []
            doc_index[doc].append(idx)
            
    # --- PRÉ-PROCESSAMENTO: AGRUPAMENTO DE SAÍDAS (MANY-TO-ONE) ---
    # Identifica grupos de saídas que devem ser somados para bater com uma entrada
    matches_agrupados = {} # Mapa: idx_saida -> match_info
    
    # Agrupa saídas por Documento + Produto
    # Só considera itens com documento válido
    df_saida_validos = df_saida[df_saida['doc_num'] != ''].copy()
    if not df_saida_validos.empty:
        # Cria chave de agrupamento
        df_saida_validos['chave_grupo'] = df_saida_validos['doc_num'] + "_" + df_saida_validos['ds_produto']
        grupos = df_saida_validos.groupby('chave_grupo')
        
        for chave, grupo in grupos:
            if len(grupo) > 1: # Só interessa se tiver mais de 1 item
                doc_grupo = grupo.iloc[0]['doc_num']
                prod_grupo = grupo.iloc[0]['ds_produto']
                comp_grupo = grupo.iloc[0]['comps']
                
                qtd_total_saida = grupo['qt_entrada'].astype(float).sum() # qt_entrada na saída é a quantidade
                valor_total_saida = grupo['valor_total'].astype(float).sum()
                
                # Busca entrada correspondente (mesmo doc, produto similar, quantidade próxima da SOMA)
                if doc_grupo in doc_index:
                    candidatos_idx = doc_index[doc_grupo]
                    
                    for idx_e in candidatos_idx:
                        if idx_e in entradas_processadas: continue
                        
                        row_e = df_entrada.loc[idx_e]
                        qtd_e = float(row_e.get('qt_entrada', 0))
                        
                        # CRITICAL: Valida qualidade do match baseado na quantidade
                        # Se quantidade bate com a SOMA, aceita threshold menor (70%)
                        # Se quantidade difere, exige threshold maior (85%)
                        qtd_match_soma = abs(qtd_e - qtd_total_saida) < 0.1
                        limiar_grupo = 70 if qtd_match_soma else 85
                        
                        # Verifica produto (ignora penalidades pois tem documento)
                        score_prod, _ = calcular_similaridade_precalc(comp_grupo, row_e['comps'], ignore_penalties=True)
                        
                        if score_prod >= limiar_grupo:
                            # Verifica se a quantidade da entrada bate com a SOMA das saídas
                            if qtd_match_soma:
                                # MATCH ENCONTRADO! (Agrupamento de Saída)
                                
                                # Marca entrada como usada
                                entradas_processadas.add(idx_e)
                                
                                # Distribui o match para cada item do grupo de saída
                                for idx_s, row_s in grupo.iterrows():
                                    qtd_s = float(row_s.get('qt_entrada', 0))
                                    perc_do_total = qtd_s / qtd_total_saida if qtd_total_saida > 0 else 0
                                    
                                    # Calcula valor proporcional da entrada
                                    valor_prop_e = float(row_e['valor_total']) * perc_do_total
                                    
                                    matches_agrupados[idx_s] = {
                                        'index': idx_e,
                                        'row': row_e,
                                        'score': 100,
                                        'score_produto': score_prod,
                                        'detalhes': f"Agrupado (Soma {len(grupo)} itens: {qtd_total_saida:.0f} un)",
                                        'detalhes_produto': "Match por agrupamento de saída",
                                        'valor_entrada_proporcional': valor_prop_e,
                                        'qtd_entrada_proporcional': qtd_s # Assume que bateu exato
                                    }
                                break # Achou match para o grupo, vai para próximo grupo
    
    stats = {
        'conformes': 0,
        'nao_conformes': 0,
        'nao_encontrados': 0,
        'valor_divergente': 0,
        'qtd_divergente': 0,
        'matches_perfeitos': 0,
        'matches_bons': 0,
        'matches_razoaveis': 0
    }
    
    total_items = len(df_saida)
    
    for idx_s, row_s in df_saida.iterrows():
        if progress_bar and idx_s % 20 == 0:  # Atualiza a cada 20 itens
            progress = 0.05 + (idx_s / total_items) * 0.9
            progress_bar.progress(progress, text=f"Analisando {idx_s + 1}/{total_items}")

        doc_num = row_s['doc_num']
        produto_s = row_s['ds_produto']
        comp_s = row_s['comps']
        valor_s = float(row_s['valor_total'])
        qtd_s = float(row_s.get('qt_entrada', 0))
        origem_s_norm = row_s['origem_norm']
        destino_s_norm = row_s['destino_norm']
        data_s = row_s['data']
        destino_eh_cp = row_s['destino_cp']
        
        # VERIFICA SE JÁ FOI RESOLVIDO POR AGRUPAMENTO
        if idx_s in matches_agrupados:
            match_info = matches_agrupados[idx_s]
            row_e = match_info['row']
            
            # Usa valores proporcionais calculados
            valor_e = match_info['valor_entrada_proporcional']
            qtd_e = match_info['qtd_entrada_proporcional']
            
            diferenca_valor = round(valor_s - valor_e, 2)
            diferenca_qtd = 0 # Considera zerado pois bateu a soma
            
            status = "✅ Conforme"
            tipo_div = "-"
            stats['conformes'] += 1
            stats['matches_perfeitos'] += 1
            
            obs = f"Score:100% | {match_info['detalhes']}"
            comp_info = f"{match_info['detalhes_produto']}"
            
            analise.append([
                data_s, row_s['unidade_origem'], row_s['unidade_destino'], doc_num, produto_s, row_e['ds_produto'],
                row_s.get('especie', ''), valor_s, valor_e, diferenca_valor, 
                qtd_s, qtd_e, diferenca_qtd,
                status, tipo_div, "⭐⭐⭐ Excelente", obs, comp_info
            ])
            continue # Pula processamento normal
        
        # ESTRATÉGIA DE BUSCA INTELIGENTE - DOCUMENTO PRIMEIRO
        candidatos_idx = []
        match_agregado = None
        candidatos = pd.DataFrame()  # Inicializa vazio por padrão
        documento_nao_encontrado = False  # Flag para prevenir fallback
        
        # PRIORIDADE 1: Verificação obrigatória do documento (exceto Casa Portugal)
        # O documento DEVE ser verificado ANTES de qualquer cruzamento por item
        if doc_num and doc_num != '': # Se tem documento na saída
            # Se tem documento, busca APENAS entradas com o mesmo documento
            if doc_num in doc_index:
                candidatos_idx = doc_index[doc_num]
                # Filtra apenas não processados para agregação
                candidatos_disponiveis = [i for i in candidatos_idx if i not in entradas_processadas]
                
                # Tenta AGREGAÇÃO (One-to-Many) se houver múltiplos candidatos do mesmo produto
                matches_doc_prod = []
                match_exato = None  # Guarda entrada com quantidade EXATA
                
                for idx_e in candidatos_disponiveis:
                    row_e = df_entrada.loc[idx_e]
                    qtd_e = float(row_e.get('qt_entrada', 0))
                    
                    # CRITICAL: Valida qualidade do match baseado na quantidade
                    # Se quantidade bate exata, aceita threshold menor (70%)
                    # Se quantidade difere, exige threshold maior (85%) para evitar matches incorretos
                    qtd_match_exato = abs(qtd_e - qtd_s) < 0.01  # Tolerância float
                    limiar_doc = 70 if qtd_match_exato else 85
                    
                    # Como estamos dentro do bloco de mesmo documento, podemos ignorar penalidades
                    score_prod, _ = calcular_similaridade_precalc(comp_s, row_e['comps'], ignore_penalties=True)
                    
                    if score_prod >= limiar_doc:
                        # PRIORIDADE: Se encontrou quantidade EXATA, guarda para usar individualmente
                        if qtd_match_exato:
                            match_exato = {
                                'index': idx_e,
                                'row': row_e,
                                'score': 100,
                                'score_produto': score_prod,
                                'detalhes': f"Match exato (Doc:{doc_num}, Qtd:{qtd_e})",
                                'detalhes_produto': "Quantidade exata"
                            }
                        matches_doc_prod.append((idx_e, row_e, score_prod))
                
                # Se encontrou match EXATO, usa ele e não agrega
                if match_exato:
                    match_agregado = match_exato
                # Senão, avalia agregação se houver múltiplos candidatos
                elif matches_doc_prod:
                    qtd_total_entrada = sum(float(m[1].get('qt_entrada', 0)) for m in matches_doc_prod)
                    qtd_primeiro = float(matches_doc_prod[0][1].get('qt_entrada', 0))
                    
                    # Desvio percentual da soma em relação à quantidade de saída
                    desvio_soma = abs(qtd_total_entrada - qtd_s) / qtd_s * 100 if qtd_s > 0 else 0
                    
                    # Só permite agregação se a soma não se afastar muito da saída (até 10%)
                    # Evita casos em que, por engano, somaria itens que deveriam casar com outras saídas
                    soma_razoavel = desvio_soma <= 10
                    
                    # Se a soma das quantidades bate melhor com a saída ou se são múltiplos itens
                    if soma_razoavel and (len(matches_doc_prod) > 1 or (abs(qtd_total_entrada - qtd_s) < abs(qtd_primeiro - qtd_s))):
                        
                        valor_total_entrada = sum(float(m[1]['valor_total']) for m in matches_doc_prod)
                        
                        # Cria um "row" virtual agregado
                        row_virtual = matches_doc_prod[0][1].copy()
                        row_virtual['qt_entrada'] = qtd_total_entrada
                        row_virtual['valor_total'] = valor_total_entrada
                        row_virtual['ds_produto'] = f"{row_virtual['ds_produto']} (+ {len(matches_doc_prod)-1} itens)"
                        
                        match_agregado = {
                            'indices': [m[0] for m in matches_doc_prod],
                            'row': row_virtual,
                            'score': 100, # Match confirmado por doc + produto
                            'score_produto': matches_doc_prod[0][2],
                            'detalhes': f"Agregado: {len(matches_doc_prod)} itens (Doc:{doc_num})",
                            'detalhes_produto': "Múltiplos itens somados"
                        }

                # Define candidatos para busca normal (caso não use o agregado ou para comparação)
                # IMPORTANTE: Só busca dentro das entradas com o mesmo documento
                candidatos = df_entrada.loc[candidatos_idx]
            else:
                # Documento não encontrado na entrada - marca como não encontrado
                # Não busca por outros critérios se o documento não existe
                candidatos = pd.DataFrame()  # DataFrame vazio
                documento_nao_encontrado = True  # CRITICAL: Previne fallback
        elif destino_eh_cp:
            # Casa Portugal: não exige documento, busca por data e produto
            if pd.notna(data_s):
                mask_data = (df_entrada['data'] >= data_s - pd.Timedelta(days=30)) & \
                            (df_entrada['data'] <= data_s + pd.Timedelta(days=30))
                candidatos = df_entrada[mask_data]
            else:
                candidatos = df_entrada
        else:
            # Sem documento válido: busca por janela de data (fallback)
            # IMPORTANTE: Só faz fallback se o documento não existia (não foi "não encontrado")
            if not documento_nao_encontrado:
                if pd.notna(data_s):
                    mask_data = (df_entrada['data'] >= data_s - pd.Timedelta(days=30)) & \
                                (df_entrada['data'] <= data_s + pd.Timedelta(days=30))
                    candidatos = df_entrada[mask_data]
                else:
                    candidatos = df_entrada
            # Se documento_nao_encontrado=True, candidatos já está vazio (linha 696)
        
        # Se ainda tem muitos candidatos, limita aos 100 mais recentes
        if len(candidatos) > 100:
            candidatos = candidatos.head(100)
        
        matches = []
        best_score = 0
        
        # Se já temos um match agregado perfeito, usamos ele
        if match_agregado:
            matches.append(match_agregado)
            best_score = 100
        else:
            # Loop normal de busca (One-to-One)
            for idx_e, row_e in candidatos.iterrows():
                if idx_e in entradas_processadas: continue # Pula já processados
                
                # Early exit
                if best_score >= 95: break
                
                score_total = 0
                detalhes_match = []
                
                # 1. DOCUMENTO PRIMEIRO (40% - prioridade máxima quando existe)
                # Verifica documento ANTES de qualquer outro critério
                doc_num_e = row_e['doc_num']
                doc_match = False
                
                if destino_eh_cp:
                    # Casa Portugal: não exige documento
                    score_total += 40
                    detalhes_match.append("Doc:CP")
                    doc_match = True
                elif doc_num and doc_num != '' and doc_num_e and doc_num_e != '':
                    # Verifica correspondência exata do documento
                    if doc_num == doc_num_e:
                        score_total += 40
                        detalhes_match.append(f"Doc:✓{doc_num}")
                        doc_match = True
                    else:
                        # Documento diferente - penalização severa
                        score_total += 0
                        detalhes_match.append(f"Doc:✗{doc_num_e}≠{doc_num}")
                        # Se o documento não corresponde, não continua (exceto se for Casa Portugal)
                        continue  # Pula este candidato se documento não corresponde
                elif not doc_num or doc_num == '':
                    # Sem documento na saída - penalização moderada
                    score_total += 15
                    detalhes_match.append("Doc:N/A(saída)")
                else:
                    # Sem documento na entrada - penalização moderada
                    score_total += 15
                    detalhes_match.append("Doc:N/A(entrada)")
                
                # 2. Produto (45% - só avalia se documento está OK ou é Casa Portugal)
                # Se tem documento correspondente, ignora penalidades de divergência leve
                score_produto, detalhes_produto = calcular_similaridade_precalc(comp_s, row_e['comps'], ignore_penalties=doc_match)
                
                # Se tem documento correspondente, valida a quantidade para definir o limiar
                if doc_match:
                    qtd_e = float(row_e.get('qt_entrada', 0))
                    qtd_match_exato = abs(qtd_e - qtd_s) < 0.01
                    # Se quantidade bate, aceita limiar baixo (40%). Se não bate, exige alto (85%)
                    limiar_efetivo = 40 if qtd_match_exato else 85
                else:
                    limiar_efetivo = limiar_similaridade
                
                if score_produto < limiar_efetivo: 
                    continue  # Produto não similar o suficiente
                
                score_total += score_produto * 0.45
                detalhes_match.append(f"Prod:{score_produto:.0f}%")
                
                # 3. Unidades (5%)
                origem_match = origem_s_norm == row_e['origem_norm']
                destino_match = destino_s_norm == row_e['destino_norm']
                
                # Verifica também o inverso (A->B vs B->A) comum em devoluções
                origem_cross = origem_s_norm == row_e['destino_norm']
                destino_cross = destino_s_norm == row_e['origem_norm']
                
                if (origem_match and destino_match) or (origem_cross and destino_cross):
                    score_total += 5
                    detalhes_match.append("Unid:✓")
                elif origem_match or destino_match or origem_cross or destino_cross:
                    score_total += 3
                    detalhes_match.append("Unid:~")
                
                # 4. Espécie (3%)
                especie_s = str(row_s.get('especie', '')).strip().upper()
                especie_e = str(row_e.get('especie', '')).strip().upper()
                if especie_s and especie_e:
                    if especie_s == especie_e:
                        score_total += 3
                        detalhes_match.append("Esp:✓")
                    else:
                        detalhes_match.append("Esp:x")
                else:
                    score_total += 2 # Se não tem espécie, não penaliza muito
                    detalhes_match.append("Esp:?")

                # 5. Data (5%)
                if pd.notna(data_s) and pd.notna(row_e['data']):
                    diff_dias = abs((row_e['data'] - data_s).days)
                    if diff_dias == 0:
                        score_total += 5
                        detalhes_match.append("Data:mesma")
                    elif diff_dias <= 3:
                        score_total += 4
                        detalhes_match.append(f"Data:{diff_dias}d")
                    elif diff_dias <= 7:
                        score_total += 3
                        detalhes_match.append(f"Data:{diff_dias}d")
                    elif diff_dias <= 15:
                        score_total += 1
                        detalhes_match.append(f"Data:{diff_dias}d")
                
                # 6. Valor (2%)
                valor_e = float(row_e['valor_total'])
                if valor_s > 0:
                    perc_diff = abs(valor_s - valor_e) / valor_s * 100
                    if perc_diff <= 1:
                        score_total += 2
                        detalhes_match.append("Valor:≈")
                    elif perc_diff <= 5:
                        score_total += 1.5
                        detalhes_match.append(f"Valor:~{perc_diff:.1f}%")
                    elif perc_diff <= 15:
                        score_total += 1
                        detalhes_match.append(f"Valor:~{perc_diff:.1f}%")
                    elif perc_diff <= 50:
                        score_total += 0.5
                        detalhes_match.append(f"Valor:diff{perc_diff:.0f}%")
                
                if score_total >= 50:
                    matches.append({
                        'index': idx_e,
                        'row': row_e,
                        'score': score_total,
                        'score_produto': score_produto,
                        'detalhes': ' | '.join(detalhes_match),
                        'detalhes_produto': detalhes_produto
                    })
                    if score_total > best_score: best_score = score_total
        
        if matches:
            matches.sort(key=lambda x: (x['score'], x['score_produto']), reverse=True)
            best_match = matches[0]
            row_e = best_match['row']
            
            # VALIDAÇÃO DE QUANTIDADE - Evita falsos positivos
            valor_e = float(row_e['valor_total'])
            qtd_e = float(row_e.get('qt_entrada', 0))
            
            # Verifica se o match de quantidade faz sentido
            doc_match = (doc_num and doc_num != '' and doc_num == row_e.get('doc_num', ''))
            is_valid, diferenca_qtd_validada = validar_match_quantidade(
                qtd_s, qtd_e, best_match['score_produto'], doc_match
            )
            
            # Se a validação falhou, trata como não encontrado
            if not is_valid:
                stats['nao_encontrados'] += 1
                stats['nao_conformes'] += 1
                analise.append([
                    data_s, row_s['unidade_origem'], row_s['unidade_destino'], doc_num, produto_s, "-",
                    row_s.get('especie', ''), valor_s, None, None, 
                    qtd_s, None, None,
                    "❌ Não Conforme", "Item não encontrado (quantidade incompatível)", "-", 
                    f"Match rejeitado: Qtd muito divergente (Saída:{qtd_s} vs Entrada:{qtd_e})", "-"
                ])
                continue  # Pula para próximo item
            
            # Match válido - prossegue com análise normal
            diferenca_qtd = diferenca_qtd_validada
            
            # Marca processados (pode ser lista de índices ou único)
            if 'indices' in best_match:
                for idx in best_match['indices']:
                    entradas_processadas.add(idx)
            else:
                entradas_processadas.add(best_match['index'])
            
            diferenca_valor = round(valor_s - valor_e, 2)
            
            perc_diff_valor = abs(diferenca_valor / valor_s * 100) if valor_s > 0 else 0
            
            if best_match['score'] >= 90:
                qualidade_match = "⭐⭐⭐ Excelente"
                stats['matches_perfeitos'] += 1
            elif best_match['score'] >= 75:
                qualidade_match = "⭐⭐ Bom"
                stats['matches_bons'] += 1
            else:
                qualidade_match = "⭐ Razoável"
                stats['matches_razoaveis'] += 1
            
            # Lógica de conformidade melhorada
            # Quantidade: tolerância de 0.01 unidades (para erros de arredondamento)
            conforme_qtd = abs(diferenca_qtd) < 0.01
            
            # Valor: tolerância inteligente baseada no valor e percentual
            # Se quantidade está igual, tolera diferenças pequenas de valor (até 10% ou R$ 1, o que for maior)
            if conforme_qtd:
                # Quantidade igual: tolera diferença percentual OU valor absoluto pequeno
                # Para valores pequenos (< R$ 10), tolera até R$ 1 de diferença
                # Para valores maiores, tolera até 10% de diferença
                if valor_s < 10:
                    # Valores pequenos: tolera até R$ 1 de diferença absoluta
                    conforme_valor = abs(diferenca_valor) <= 1.0
                else:
                    # Valores maiores: tolera até 10% de diferença OU R$ 10 fixo
                    limite_valor_absoluto = max(10.0, valor_s * 0.10)
                    conforme_valor = abs(diferenca_valor) <= limite_valor_absoluto or perc_diff_valor <= 10
            else:
                # Quantidade diferente: usa tolerância fixa de R$ 10
                conforme_valor = abs(diferenca_valor) <= 10
            
            if conforme_valor and conforme_qtd:
                status = "✅ Conforme"
                tipo_div = "-"
                stats['conformes'] += 1
            else:
                status = "⚠️ Não Conforme"
                stats['nao_conformes'] += 1
                
                tipos_div = []
                if not conforme_valor:
                    stats['valor_divergente'] += 1
                    if perc_diff_valor > 50:
                        tipos_div.append("Valor muito divergente (>50%)")
                    elif perc_diff_valor > 20:
                        tipos_div.append("Valor divergente (>20%)")
                    else:
                        tipos_div.append("Pequena divergência valor")
                
                if not conforme_qtd:
                    stats['qtd_divergente'] += 1
                    tipos_div.append(f"Divergência Qtd ({diferenca_qtd:+g})")
                
                tipo_div = " | ".join(tipos_div)
            
            obs = f"Score:{best_match['score']:.0f}% | {best_match['detalhes']}"
            comp_info = f"{best_match['detalhes_produto']}"
            
            analise.append([
                data_s, row_s['unidade_origem'], row_s['unidade_destino'], doc_num, produto_s, row_e['ds_produto'],
                row_s.get('especie', ''), valor_s, valor_e, diferenca_valor, 
                qtd_s, qtd_e, diferenca_qtd,
                status, tipo_div, qualidade_match, obs, comp_info
            ])
        else:
            stats['nao_encontrados'] += 1
            stats['nao_conformes'] += 1
            
            # Determina a razão do não encontrado
            if doc_num and doc_num != '' and not destino_eh_cp:
                motivo = f"Documento {doc_num} não encontrado na entrada"
            elif destino_eh_cp:
                motivo = "Item não encontrado (Casa Portugal)"
            else:
                motivo = "Item não encontrado na entrada (sem documento para validação)"
            
            analise.append([
                data_s, row_s['unidade_origem'], row_s['unidade_destino'], doc_num, produto_s, "-",
                row_s.get('especie', ''), valor_s, None, None, 
                qtd_s, None, None,
                "❌ Não Conforme", motivo, "-", "Sem correspondência encontrada", "-"
            ])
    
    if progress_bar:
        progress_bar.progress(0.95, text="Finalizando...")
            
    # Entradas órfãs
    for idx_e, row_e in df_entrada.iterrows():
        if idx_e in entradas_processadas:
            continue
        
        # Ignora entradas fora do período analisado
        data_e = row_e['data']
        if pd.notna(periodo_inicio) and pd.notna(periodo_fim) and pd.notna(data_e):
            if data_e < periodo_inicio or data_e > periodo_fim:
                continue
        
        doc_num_e = row_e['doc_num']
        produto_e = row_e['ds_produto']
        qtd_e = float(row_e.get('qt_entrada', 0))
        
        analise.append([
            row_e['data'], row_e['unidade_origem'], row_e['unidade_destino'],
            doc_num_e, "-", produto_e, row_e.get('especie', ''),
            None, float(row_e['valor_total']), None, 
            None, qtd_e, None,
            "❌ Não Conforme", "Item recebido sem saída correspondente", "-",
            "Entrada órfã", "-"
        ])
    
    df_resultado = pd.DataFrame(analise, columns=[
        "Data", "Unidade Origem", "Unidade Destino", "Documento",
        "Produto (Saída)", "Produto (Entrada)", "Espécie", 
        "Valor Saída (R$)", "Valor Entrada (R$)", "Diferença (R$)",
        "Qtd Saída", "Qtd Entrada", "Diferença Qtd",
        "Status", "Tipo de Divergência", 
        "Qualidade Match", "Observações", "Detalhes Produto"
    ])
    
    if progress_bar:
        progress_bar.progress(1.0, text="Concluído!")
    
    return df_resultado, stats

def gerar_excel_bytes(df):
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
            # Converte coluna de data de volta para datetime
            if 'Data' in df.columns:
                df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
            
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

col_logo, col_title, col_opts = st.columns([1, 4, 1])

with col_opts:
    if st.button("🔄 Reiniciar", use_container_width=True, type="secondary", help="Limpa a análise e anexo atual"):
        st.session_state.df_resultado = None
        st.session_state.current_metadata = None
        st.rerun()

with col_logo:
    try:
        st.image("logo.png", use_container_width=True)
    except:
        st.warning("Logo?")

with col_title:
    st.title("Dashboard de Análise")



# --- CONFIGURAÇÃO E UPLOADS (EXPANDER) ---
# Expandido se NÃO tiver resultado ainda
expander_open = st.session_state.df_resultado is None

if expander_open:
    st.markdown("""
    ### Instruções:
    1.  Arraste os arquivos de **Saída** e **Entrada** para a área de upload.
    2.  O sistema identificará automaticamente qual é qual.
    3.  Ajuste os filtros se necessário e explore os resultados!
    """)
with st.expander("📁 Upload e Configurações da Análise", expanded=expander_open):
    col_up, col_param = st.columns([2, 1])
    
    with col_up:
        uploaded_files = st.file_uploader(
            "Anexar Arquivos (XLSX/XLS)", 
            type=["xlsx", "xls"], 
            accept_multiple_files=True,
            label_visibility="collapsed"
        )
    
    with col_param:
        limiar = 65 # Valor padrão fixo
        
        if uploaded_files and len(uploaded_files) >= 2:
            processar = st.button("🚀 Processar Análise", type="primary", use_container_width=True)
        else:
            if not uploaded_files:
                st.caption("Anexe arquivos para começar.")
            elif len(uploaded_files) < 2:
                st.caption("Mínimo 2 arquivos.")
            processar = False



# --- Lógica de Processamento ---
if 'df_resultado' not in st.session_state:
    st.session_state.df_resultado = None
if 'current_metadata' not in st.session_state:
    st.session_state.current_metadata = None

if processar and uploaded_files:
    with st.spinner("Processando arquivos..."):
        # Classifica todos os arquivos em Saída ou Entrada
        arquivos_saida = []
        arquivos_entrada = []
        
        for file in uploaded_files:
            df_temp = pd.read_excel(file)
            
            # Pontua o arquivo
            def pontuar_arquivo(df, nome_arquivo, termos_saida, termos_entrada):
                score_saida = sum(1 for t in termos_saida if t in nome_arquivo.lower())
                score_entrada = sum(1 for t in termos_entrada if t in nome_arquivo.lower())
                return score_saida, score_entrada
            
            s_saida, s_entrada = pontuar_arquivo(df_temp, file.name, ['saida', 'concedido', 'envio'], ['entrada', 'recebido'])
            
            if s_saida > s_entrada:
                arquivos_saida.append((file, df_temp))
            else:
                arquivos_entrada.append((file, df_temp))
        
        # Verifica se tem pelo menos 1 de cada tipo
        if not arquivos_saida or not arquivos_entrada:
            st.error("❌ É necessário ter pelo menos 1 arquivo de Saída e 1 de Entrada!")
            st.stop()
        
        # Consolida múltiplos arquivos do mesmo tipo
        st.toast(f"📤 {len(arquivos_saida)} arquivo(s) de Saída identificado(s)")
        st.toast(f"📥 {len(arquivos_entrada)} arquivo(s) de Entrada identificado(s)")
        
        # Mescla arquivos de Saída
        dfs_saida_list = []
        nomes_saida = []
        for file, df in arquivos_saida:
            df.columns = [c.strip().lower() for c in df.columns]
            dfs_saida_list.append(df)
            nomes_saida.append(file.name)
        
        df_saida = pd.concat(dfs_saida_list, ignore_index=True) if len(dfs_saida_list) > 1 else dfs_saida_list[0]
        nome_saida = f"{len(nomes_saida)} arquivo(s) consolidado(s)" if len(nomes_saida) > 1 else nomes_saida[0]
        
        # Mescla arquivos de Entrada
        dfs_entrada_list = []
        nomes_entrada = []
        for file, df in arquivos_entrada:
            df.columns = [c.strip().lower() for c in df.columns]
            dfs_entrada_list.append(df)
            nomes_entrada.append(file.name)
        
        df_entrada = pd.concat(dfs_entrada_list, ignore_index=True) if len(dfs_entrada_list) > 1 else dfs_entrada_list[0]
        nome_entrada = f"{len(nomes_entrada)} arquivo(s) consolidado(s)" if len(nomes_entrada) > 1 else nomes_entrada[0]
        
        # Mapeia colunas para os nomes esperados pela função analisar_itens
        def mapear_colunas(df):
            mapeamento = {}
            for col in df.columns:
                col_lower = col.lower()
                # Produto/Descrição
                if any(x in col_lower for x in ['produto', 'descrição', 'descricao', 'material']):
                    if 'ds_produto' not in mapeamento.values():
                        mapeamento[col] = 'ds_produto'
                # Documento
                elif any(x in col_lower for x in ['documento', 'nf', 'nota']):
                    if 'documento' not in mapeamento.values():
                        mapeamento[col] = 'documento'
                # Unidade Origem
                elif 'origem' in col_lower and 'unidade' in col_lower:
                    mapeamento[col] = 'unidade_origem'
                # Unidade Destino
                elif 'destino' in col_lower and 'unidade' in col_lower:
                    mapeamento[col] = 'unidade_destino'
                # Valor Total
                elif any(x in col_lower for x in ['valor total', 'vl_total', 'total']):
                    if 'valor_total' not in mapeamento.values():
                        mapeamento[col] = 'valor_total'
                # Quantidade
                elif any(x in col_lower for x in ['quantidade', 'qtd', 'qt_entrada', 'qt entrada']):
                    if 'qt_entrada' not in mapeamento.values():
                        mapeamento[col] = 'qt_entrada'
                # Espécie
                elif 'especie' in col_lower or 'espécie' in col_lower:
                    if 'especie' not in mapeamento.values():
                        mapeamento[col] = 'especie'
            return mapeamento
        
        # Aplica mapeamento
        map_saida = mapear_colunas(df_saida)
        map_entrada = mapear_colunas(df_entrada)
        
        # Debug: mostra colunas mapeadas (descomente se precisar debugar)
        # st.toast(f"Colunas Saída: {list(df_saida.columns)}")
        # st.toast(f"Mapeamento Saída: {map_saida}")
        
        df_saida.rename(columns=map_saida, inplace=True)
        df_entrada.rename(columns=map_entrada, inplace=True)
        
        # Verifica se todas as colunas necessárias existem
        colunas_necessarias = ['ds_produto', 'documento', 'unidade_origem', 'unidade_destino', 'valor_total', 'data']
        for col in colunas_necessarias:
            if col not in df_saida.columns:
                st.error(f"❌ Coluna '{col}' não encontrada em Saída. Colunas disponíveis: {list(df_saida.columns)}")
            if col not in df_entrada.columns:
                st.error(f"❌ Coluna '{col}' não encontrada em Entrada. Colunas disponíveis: {list(df_entrada.columns)}")
        
        df_saida['data'] = _parse_date_column(df_saida['data'])
        df_entrada['data'] = _parse_date_column(df_entrada['data'])
        
        # Normaliza valores numéricos (remove separadores de milhar)
        if 'qt_entrada' in df_saida.columns:
            df_saida['qt_entrada'] = df_saida['qt_entrada'].apply(normalizar_valor_numerico)
        if 'qt_entrada' in df_entrada.columns:
            df_entrada['qt_entrada'] = df_entrada['qt_entrada'].apply(normalizar_valor_numerico)
        if 'valor_total' in df_saida.columns:
            df_saida['valor_total'] = df_saida['valor_total'].apply(normalizar_valor_numerico)
        if 'valor_total' in df_entrada.columns:
            df_entrada['valor_total'] = df_entrada['valor_total'].apply(normalizar_valor_numerico)
        
        # Processa
        progress_bar = st.progress(0, text="Iniciando...")
        df_res, stats = analisar_itens(df_saida, df_entrada, limiar, progress_bar)
        progress_bar.empty()
        
        # Salva no histórico
        analysis_id = save_analysis_to_history(df_res, stats, nome_saida, nome_entrada)
        if analysis_id:
            st.toast("Análise salva no histórico!")
        
        st.session_state.df_resultado = df_res
        st.session_state.current_metadata = {
            'arquivo_saida': nome_saida,
            'arquivo_entrada': nome_entrada,
            'data_formatada': datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        }
        st.toast("✅ Análise concluída!", icon="✅")

# --- Dashboard ---
if st.session_state.df_resultado is not None:
    df = st.session_state.df_resultado
    
    # Mostra informações da análise atual
    # Mostra informações do período apurado
    min_date_apurado = df['Data'].min()
    max_date_apurado = df['Data'].max()
    
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
            min_date = df['Data'].min()
            max_date = df['Data'].max()
            date_range = st.date_input("Período", [min_date, max_date])
    
    # Aplica Filtros
    df_filtered = df[df['Status'].isin(status_filter)]
    
    if unidade_filter:
        df_filtered = df_filtered[
            df_filtered['Unidade Origem'].isin(unidade_filter) | 
            df_filtered['Unidade Destino'].isin(unidade_filter)
        ]
        
    if len(date_range) == 2:
        df_filtered = df_filtered[
            (df_filtered['Data'].dt.date >= date_range[0]) & 
            (df_filtered['Data'].dt.date <= date_range[1])
        ]
    
    # --- Balanço Financeiro do Período ---
    st.markdown("### Balanço Financeiro do Período")
    
    total_saida_periodo = df_filtered['Valor Saída (R$)'].sum()
    total_entrada_periodo = df_filtered['Valor Entrada (R$)'].sum()
    diff_financeira_periodo = total_saida_periodo - total_entrada_periodo
    
    col_balanco1, col_balanco2, col_balanco3 = st.columns(3)
    
    with col_balanco1:
        st.metric("Total Saída (Enviado)", f"R$ {total_saida_periodo:_.2f}".replace('.', ',').replace('_', '.'), help="Soma do valor de todos os itens de saída no período")
        
    with col_balanco2:
        st.metric("Total Entrada (Recebido)", f"R$ {total_entrada_periodo:_.2f}".replace('.', ',').replace('_', '.'), help="Soma do valor de todos os itens de entrada no período")
        
    with col_balanco3:
        cor_diff = "normal"
        if diff_financeira_periodo > 0:
            cor_diff = "inverse" # Vermelho/Negativo (saiu mais que entrou)
        elif diff_financeira_periodo < 0:
            cor_diff = "off" # Verde/Positivo (entrou mais que saiu - raro mas possível)
            
        st.metric(
            "Divergência Financeira", 
            f"R$ {diff_financeira_periodo:_.2f}".replace('.', ',').replace('_', '.'), 
            delta=f"{(diff_financeira_periodo/total_saida_periodo*100):.1f}%" if total_saida_periodo > 0 else "0%",
            delta_color="inverse",
            help="Balanço Líquido: Total Saída - Total Entrada. Indica se 'sobrou' ou 'faltou' valor no total geral."
        )
    
    st.divider()

    # --- KPIs Premium ---
    st.markdown("### Indicadores Principais")
    
    # CSS para valores divergentes em vermelho (ambos os temas) e ajuste de fonte
    st.markdown("""
        <style>
        /* Ajuste de tamanho de fonte para caber 5 colunas */
        [data-testid="stMetricValue"] {
            font-size: 22px !important;
        }
        
        /* KPI Não Conformes - texto vermelho - mais específico */
        div[data-testid="stMetric"]:nth-of-type(3) [data-testid="stMetricValue"],
        div[data-testid="stMetric"]:nth-of-type(3) [data-testid="stMetricValue"] div,
        div[data-testid="stMetric"]:nth-of-type(3) [data-testid="stMetricValue"] * {
            color: #FF4444 !important;
        }
        /* KPI Valor Divergente - texto vermelho - mais específico */
        div[data-testid="stMetric"]:nth-of-type(4) [data-testid="stMetricValue"],
        div[data-testid="stMetric"]:nth-of-type(4) [data-testid="stMetricValue"] div,
        div[data-testid="stMetric"]:nth-of-type(4) [data-testid="stMetricValue"] * {
            color: #FF4444 !important;
        }
        /* KPI Divergência Qtd - texto vermelho - mais específico */
        div[data-testid="stMetric"]:nth-of-type(5) [data-testid="stMetricValue"],
        div[data-testid="stMetric"]:nth-of-type(5) [data-testid="stMetricValue"] div,
        div[data-testid="stMetric"]:nth-of-type(5) [data-testid="stMetricValue"] * {
            color: #FF4444 !important;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # Ajusta largura das colunas para dar mais espaço ao Valor Divergente
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns([1, 1, 1, 1.3, 1])
    
    total_analisado = len(df_filtered)
    
    total_conforme = len(df_filtered[df_filtered['Status'].str.contains('Conforme') & ~df_filtered['Status'].str.contains('Não')])
    total_nao_conforme = len(df_filtered[df_filtered['Status'].str.contains('Não Conforme')])
    
    # Cálculos de divergência
    valor_divergente = df_filtered[df_filtered['Status'].str.contains('Não Conforme')]['Diferença (R$)'].abs().sum()
    
    # Conta itens com divergência de quantidade (excluindo não encontrados/nulos)
    # Considera divergência se 'Diferença Qtd' não é nulo e é diferente de 0
    qtd_divergente_count = len(df_filtered[
        (df_filtered['Diferença Qtd'].notna()) & 
        (df_filtered['Diferença Qtd'] != 0)
    ])
    
    # Calcula percentuais
    perc_conforme = (total_conforme / total_analisado * 100) if total_analisado > 0 else 0
    perc_nao_conforme = (total_nao_conforme / total_analisado * 100) if total_analisado > 0 else 0
    
    # Formata valor com ponto para milhar e vírgula para decimal
    def formatar_valor(valor):
        # Formata com 2 casas decimais
        valor_str = f"{valor:_.2f}".replace("_", ".")
        # Separa parte inteira e decimal
        partes = valor_str.split(".")
        if len(partes) == 3:  # tem milhar
            return f"{partes[0]}.{partes[1]},{partes[2]}"
        else:  # sem milhar
            return f"{partes[0]},{partes[1]}"
    
    with kpi1:
        st.metric("Total de Itens", f"{total_analisado:_.0f}".replace("_", "."), help="Total de registros analisados")
    with kpi2:
        st.metric("Conformes", f"{total_conforme:_.0f}".replace("_", "."), delta=f"{perc_conforme:.1f}%", help="Itens em conformidade")
    with kpi3:
        st.metric("Não Conformes", f"{total_nao_conforme:_.0f}".replace("_", "."), delta=f"-{perc_nao_conforme:.1f}%", delta_color="inverse", help="Itens com divergências")
    with kpi4:
        st.metric("Valor Divergente", f"R$ {formatar_valor(valor_divergente)}", help="Soma ABSOLUTA das diferenças dos itens Não Conformes. Ex: Falta 10 + Sobra 10 = 20 de divergência.")
    with kpi5:
        st.metric("Divergência Qtd", f"{qtd_divergente_count:_.0f}".replace("_", "."), help="Itens com diferença na quantidade (Saída vs Entrada)")
    
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
    
    # Define cor do texto dos gráficos (sempre claro)
    chart_text_color = '#001A72'
    chart_grid_color = 'rgba(128,128,128,0.2)'

    # --- Gráficos Premium ---
    col_chart1, col_chart2 = st.columns(2)
    
    with col_chart1:
        st.markdown("#### Distribuição de Status")
        
        # Conta status
        status_counts = df_filtered['Status'].value_counts()
        
        # Gráfico de rosca com Plotly - texto otimizado
        fig_status = go.Figure(data=[go.Pie(
            labels=status_counts.index,
            values=status_counts.values,
            hole=0.6,
            marker=dict(
                colors=['#00C853', '#FF6B6B', '#FFA726'],
                line=dict(color='white', width=3)
            ),
            textposition='outside',
            textinfo='label+percent',
            textfont=dict(size=13, family="Arial", color=chart_text_color),
            insidetextorientation='radial',
            pull=[0.05, 0.05, 0.05],  # Separa levemente as fatias
            hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>'
        )])
        
        fig_status.update_layout(
            showlegend=False,  # Remove legenda duplicada
            height=380,
            margin=dict(t=10, b=10, l=10, r=10),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Arial", size=13, color=chart_text_color)
        )
        
        st.plotly_chart(fig_status, use_container_width=True)
        
    with col_chart2:
        st.markdown("#### Top 5 Divergências")
        
        df_div = df_filtered[df_filtered['Status'].str.contains('Não Conforme')]
        if not df_div.empty:
            div_counts = df_div['Tipo de Divergência'].value_counts().head(5)
            
            fig_div = go.Figure(data=[go.Bar(
                x=div_counts.values,
                y=div_counts.index,
                orientation='h',
                marker=dict(
                    color=div_counts.values,
                    colorscale='Reds',
                    line=dict(color='white', width=1)
                ),
                text=div_counts.values,
                textposition='outside',
                textfont=dict(color=chart_text_color),
                hovertemplate='<b>%{y}</b><br>Quantidade: %{x}<extra></extra>'
            )])
            
            # Calcula limite do eixo X com folga para o texto
            max_val = div_counts.values.max()
            
            fig_div.update_layout(
                height=350,
                margin=dict(t=20, b=20, l=20, r=50),  # Margem direita aumentada
                xaxis_title="Quantidade",
                yaxis_title="",
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(family="Arial", size=12, color=chart_text_color),
                xaxis=dict(
                    showgrid=True, 
                    gridcolor=chart_grid_color, 
                    title_font=dict(color=chart_text_color), 
                    tickfont=dict(color=chart_text_color),
                    range=[0, max_val * 1.2]  # 20% de folga
                ),
                yaxis=dict(showgrid=False, tickfont=dict(color=chart_text_color))
            )
            
            st.plotly_chart(fig_div, use_container_width=True)
        else:
            st.info("Nenhuma divergência encontrada!")
    
    # --- Linha 2: Análise Temporal e Hospitais ---
    col_chart3, col_chart4 = st.columns(2)
    
    with col_chart3:
        st.markdown("#### Evolução Temporal")
        
        # Agrupa por data
        df_temp = df_filtered.copy()
        df_temp['Data_Agrupada'] = pd.to_datetime(df_temp['Data']).dt.date
        temporal = df_temp.groupby(['Data_Agrupada', 'Status']).size().reset_index(name='Quantidade')
        
        fig_temporal = px.line(
            temporal,
            x='Data_Agrupada',
            y='Quantidade',
            color='Status',
            markers=True,
            color_discrete_map={
                '✅ Conforme': '#00C853',
                '⚠️ Não Conforme': '#FF6B6B',
                '❌ Não Conforme': '#FFA726'
            }
        )
        
        fig_temporal.update_layout(
            height=350,
            margin=dict(t=20, b=60, l=20, r=20),
            xaxis_title="Data",
            yaxis_title="Quantidade",
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Arial", size=12, color=chart_text_color),
            hovermode='x unified',
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.4,
                xanchor="center",
                x=0.5,
                font=dict(color=chart_text_color)
            ),
            xaxis=dict(
                tickangle=-45,  # Rotaciona labels
                tickmode='auto',
                nticks=10,  # Limita número de ticks
                tickfont=dict(size=10, color=chart_text_color),
                title_font=dict(color=chart_text_color),
                gridcolor=chart_grid_color
            ),
            yaxis=dict(
                tickfont=dict(color=chart_text_color),
                title_font=dict(color=chart_text_color),
                gridcolor=chart_grid_color
            )
        )
        
        st.plotly_chart(fig_temporal, use_container_width=True)
    
    with col_chart4:
        st.markdown("#### Top 5 Hospitais (Divergências)")
        
        df_div = df_filtered[df_filtered['Status'].str.contains('Não Conforme')]
        if not df_div.empty:
            # Combina origem e destino
            hospitais_div = pd.concat([
                df_div['Unidade Origem'].value_counts(),
                df_div['Unidade Destino'].value_counts()
            ]).groupby(level=0).sum().sort_values(ascending=False).head(5)
            
            fig_hosp = go.Figure(data=[go.Bar(
                x=hospitais_div.values,
                y=hospitais_div.index,
                orientation='h',
                marker=dict(
                    color='#E87722',
                    line=dict(color='white', width=1)
                ),
                text=hospitais_div.values,
                textposition='outside',
                textfont=dict(color=chart_text_color),
                hovertemplate='<b>%{y}</b><br>Divergências: %{x}<extra></extra>'
            )])
            
            # Calcula limite do eixo X com folga
            max_val_hosp = hospitais_div.values.max()
            
            fig_hosp.update_layout(
                height=350,
                margin=dict(t=20, b=20, l=20, r=50),  # Margem direita aumentada
                xaxis_title="Quantidade de Divergências",
                yaxis_title="",
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(family="Arial", size=12, color=chart_text_color),
                xaxis=dict(
                    showgrid=True, 
                    gridcolor=chart_grid_color, 
                    title_font=dict(color=chart_text_color), 
                    tickfont=dict(color=chart_text_color),
                    range=[0, max_val_hosp * 1.2]  # 20% de folga
                ),
                yaxis=dict(showgrid=False, tickfont=dict(color=chart_text_color))
            )
            
            st.plotly_chart(fig_hosp, use_container_width=True)
        else:
            st.info("Nenhuma divergência por hospital!")
    
    st.divider()

    # --- Tabela Detalhada ---
    st.subheader("Detalhamento dos Dados")
    
    st.dataframe(
        df_filtered,
        use_container_width=True,
        column_config={
            "Data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
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

