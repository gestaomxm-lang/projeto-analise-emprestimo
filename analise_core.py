import pandas as pd
import re
from difflib import SequenceMatcher
from datetime import datetime
import numpy as np

# --- Funções Auxiliares de Tratamento de Dados ---

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
    """Normaliza valores numéricos removendo separadores de milhar (ponto) e convertendo para float."""
    if pd.isna(valor) or valor == '' or valor is None:
        return 0.0
    
    if isinstance(valor, (int, float)):
        return float(valor)
    
    valor_str = str(valor).strip().replace(' ', '')
    
    if not valor_str:
        return 0.0
    
    if ',' in valor_str and '.' in valor_str:
        valor_str = valor_str.replace('.', '').replace(',', '.')
    elif ',' in valor_str:
        valor_str = valor_str.replace(',', '.')
    elif '.' in valor_str:
        partes = valor_str.split('.')
        if len(partes) > 2:
            valor_str = valor_str.replace('.', '')
        elif len(partes) == 2:
            parte_decimal = partes[1]
            if parte_decimal == '000' or len(parte_decimal) > 3:
                valor_str = valor_str.replace('.', '')
    
    try:
        return float(valor_str)
    except (ValueError, TypeError):
        return 0.0

def normalizar_unidade_medida(texto):
    """Normaliza unidades de medida para formato padrão."""
    mapeamento = {
        'GR': 'G', 'GRAMA': 'G', 'GRAMAS': 'G',
        'MILIGRAMA': 'MG', 'MILIGRAMAS': 'MG',
        'MILILITRO': 'ML', 'MILILITROS': 'ML',
        'MICROGRAMA': 'MCG', 'MICROGRAMAS': 'MCG',
        'UNIDADE': 'UI', 'UNIDADES': 'UI',
        'LITRO': 'L', 'LITROS': 'L',
        'METRO': 'M', 'METROS': 'M',
        'CENTIMETRO': 'CM', 'CENTIMETROS': 'CM',
        'MILIMETRO': 'MM', 'MILIMETROS': 'MM'
    }
    
    texto_norm = texto.upper()
    for variacao, padrao in mapeamento.items():
        texto_norm = re.sub(r'\b' + variacao + r'\b', padrao, texto_norm)
    
    return texto_norm

def normalizar_dimensao(dimensao_str):
    """Normaliza dimensões para comparação consistente."""
    dimensao_str = re.sub(r'\s+', '', dimensao_str.upper())
    numeros = re.findall(r'\d+\.?\d*', dimensao_str)
    numeros_norm = [str(float(n)) for n in numeros if n]
    return 'X'.join(sorted(numeros_norm))

def extrair_e_normalizar_concentracao(descricao):
    """Extrai concentração e normaliza para formato comparável."""
    descricao = normalizar_unidade_medida(descricao)
    padroes = [
        r'(\d+[,.]?\d*)\s*(MG|G|ML|MCG|UI|L|%)\s*/\s*(\d+[,.]?\d*)\s*(MG|G|ML|MCG|UI|L)',
        r'(\d+[,.]?\d*)\s*(MG|G|ML|MCG|UI|L|%)(?!\s*/)',
    ]
    concentracoes = []
    for padrao in padroes:
        matches = re.findall(padrao, descricao)
        for match in matches:
            if isinstance(match, tuple):
                conc_str = ''.join(str(c) for c in match).replace(',', '.')
                concentracoes.append(conc_str)
    return ' '.join(concentracoes) if concentracoes else ''

def extrair_componentes_produto(descricao):
    """Extrai componentes principais do produto para matching inteligente."""
    descricao = str(descricao).upper().strip()
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
    
    texto_limpo = re.sub(r'[^\w\s]', ' ', descricao)
    texto_limpo = re.sub(r'\s+', ' ', texto_limpo).strip()
    componentes['normalizado'] = texto_limpo
    componentes['concentracao'] = extrair_e_normalizar_concentracao(descricao)
    
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
            
    qtd_match = re.search(r'(?:C/|C |X|COM )\s*(\d+)', descricao)
    if qtd_match:
        componentes['quantidade'] = qtd_match.group(1)
        
    dimensoes = re.search(r'\d+\.?\d*\s*[xX]\s*\d+\.?\d*', descricao)
    if dimensoes:
        componentes['dimensao'] = normalizar_dimensao(dimensoes.group())
    
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
    
    # Sinônimos
    sinonimos = {
        'AVENTAL': ['CAPOTE', 'AVENTAL', 'JALECO'],
        'CAPOTE': ['AVENTAL', 'CAPOTE', 'JALECO'],
        'JALECO': ['AVENTAL', 'CAPOTE', 'JALECO'],
        'ALGODAO': ['POLYCOT', 'ALGODAO', 'COTTON'],
        'POLYCOT': ['ALGODAO', 'POLYCOT', 'COTTON'],
        'COTTON': ['ALGODAO', 'POLYCOT', 'COTTON'],
        'GAZE': ['COMPRESSA', 'GAZE'],
        'COMPRESSA': ['GAZE', 'COMPRESSA'],
        'SORO': ['SOLUCAO', 'SORO', 'SOL'],
        'SOLUCAO': ['SORO', 'SOLUCAO', 'SOL'],
        'SOL': ['SORO', 'SOLUCAO', 'SOL'],
        'SALINA': ['NACL', 'CLORETO', 'SALINA', 'SF'],
        'NACL': ['SALINA', 'CLORETO', 'NACL', 'SF'],
        'SF': ['SALINA', 'NACL', 'CLORETO', 'SF'],
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
        'INJETAVEL': ['INJ', 'INJETAVEL', 'IV', 'IM', 'SC'],
        'INJ': ['INJETAVEL', 'INJ', 'IV', 'IM', 'SC'],
        'ORAL': ['VO', 'ORAL', 'BUCAL'],
        'VO': ['ORAL', 'VO', 'BUCAL'],
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
    
    sim_geral = SequenceMatcher(None, comp1['normalizado'], comp2['normalizado']).ratio()
    score += sim_geral * 30
    detalhes.append(f"Texto:{sim_geral:.0%}")
    
    if comp1['principio_ativo'] and comp2['principio_ativo']:
        sim_principio = SequenceMatcher(None, comp1['principio_ativo'], comp2['principio_ativo']).ratio()
        score += sim_principio * 35
        detalhes.append(f"Princípio:{sim_principio:.0%}")
    
    if comp1['concentracao'] and comp2['concentracao']:
        c1 = comp1['concentracao'].replace(' ', '').upper()
        c2 = comp2['concentracao'].replace(' ', '').upper()
        if c1 == c2:
            score += 20
            detalhes.append(f"Conc:✓")
        else:
            nums1 = re.findall(r'\d+\.?\d*', c1)
            nums2 = re.findall(r'\d+\.?\d*', c2)
            nums_comum = set(nums1) & set(nums2)
            if nums_comum and len(nums_comum) >= len(nums1) * 0.5:
                score += 15
                detalhes.append(f"Conc:~")
            else:
                sim_conc = SequenceMatcher(None, c1, c2).ratio()
                if sim_conc > 0.7:
                    score += sim_conc * 15
                    detalhes.append(f"Conc:~{sim_conc:.0%}")
                elif not ignore_penalties:
                    score -= 25
                    detalhes.append(f"Conc:Mismatch")
    
    if 'dimensao' in comp1 and 'dimensao' in comp2 and comp1['dimensao'] and comp2['dimensao']:
        d1_norm = comp1['dimensao']
        d2_norm = comp2['dimensao']
        if d1_norm == d2_norm:
            score += 15
            detalhes.append(f"Dim:✓")
        else:
            nums1 = set(d1_norm.split('X'))
            nums2 = set(d2_norm.split('X'))
            comum = nums1 & nums2
            if len(comum) >= 2:
                score += 10
                detalhes.append(f"Dim:~")
            elif len(comum) >= 1:
                score += 5
                detalhes.append(f"Dim:part")
            elif not ignore_penalties:
                score -= 15
                detalhes.append(f"Dim:Mismatch")

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
                score -= 10
                detalhes.append(f"Apres:Mismatch")

    palavras_comum = set(comp1['palavras_chave']) & set(comp2['palavras_chave'])
    if palavras_comum:
        perc_comum = len(palavras_comum) / max(len(comp1['palavras_chave']), len(comp2['palavras_chave']))
        score += perc_comum * 5
        detalhes.append(f"Palavras:{len(palavras_comum)}")
    
    return score, ' | '.join(detalhes)

def validar_match_quantidade(qtd_saida, qtd_entrada, score_produto, doc_match):
    """Valida se o match de quantidade faz sentido."""
    if abs(qtd_saida - qtd_entrada) < 0.01:
        return True, 0.0
    
    qtd_max = max(qtd_saida, qtd_entrada)
    if qtd_max == 0:
        return True, 0.0
    
    perc_diff = abs(qtd_saida - qtd_entrada) / qtd_max * 100
    
    if doc_match and score_produto >= 85:
        if perc_diff <= 20:
            return True, qtd_saida - qtd_entrada
    
    if not doc_match:
        if perc_diff > 10:
            return False, None
    
    return True, qtd_saida - qtd_entrada

def eh_casa_portugal(unidade):
    unidade_norm = str(unidade).upper().strip()
    return 'CASA' in unidade_norm and 'PORTUGAL' in unidade_norm

def analisar_itens(df_saida, df_entrada, limiar_similaridade=65, progress_callback=None):
    """
    Executa a análise entre dataframes de saída e entrada.
    progress_callback: função que recebe (float, str) para reportar progresso.
    """
    analise = []
    entradas_processadas = set()
    
    periodo_inicio = df_saida['data'].min() if 'data' in df_saida.columns else None
    periodo_fim = df_saida['data'].max() if 'data' in df_saida.columns else None
    
    # Normalização
    for df in [df_saida, df_entrada]:
        df['documento'] = df['documento'].astype(str).str.strip()
        df['ds_produto'] = df['ds_produto'].astype(str).str.strip()
        df['unidade_origem'] = df['unidade_origem'].astype(str).str.strip()
        df['unidade_destino'] = df['unidade_destino'].astype(str).str.strip()
    
    if progress_callback:
        progress_callback(0.05, "Pré-processando dados...")
    
    # Pré-cálculos
    df_saida['comps'] = df_saida['ds_produto'].apply(extrair_componentes_produto)
    df_entrada['comps'] = df_entrada['ds_produto'].apply(extrair_componentes_produto)
    
    df_saida['doc_num'] = df_saida['documento'].apply(extrair_numeros)
    df_entrada['doc_num'] = df_entrada['documento'].apply(extrair_numeros)
    
    df_saida['destino_cp'] = df_saida['unidade_destino'].apply(eh_casa_portugal)
    
    df_saida['origem_norm'] = df_saida['unidade_origem'].str.upper().str.strip()
    df_saida['destino_norm'] = df_saida['unidade_destino'].str.upper().str.strip()
    df_entrada['origem_norm'] = df_entrada['unidade_origem'].str.upper().str.strip()
    df_entrada['destino_norm'] = df_entrada['unidade_destino'].str.upper().str.strip()
    
    # Índice
    doc_index = {}
    for idx, row in df_entrada.iterrows():
        doc = row['doc_num']
        if doc and doc != '':
            if doc not in doc_index:
                doc_index[doc] = []
            doc_index[doc].append(idx)
            
    matches_agrupados = {}
    
    # Agrupamento
    df_saida_validos = df_saida[df_saida['doc_num'] != ''].copy()
    if not df_saida_validos.empty:
        df_saida_validos['chave_grupo'] = df_saida_validos['doc_num'] + "_" + df_saida_validos['ds_produto']
        grupos = df_saida_validos.groupby('chave_grupo')
        
        for chave, grupo in grupos:
            if len(grupo) > 1:
                doc_grupo = grupo.iloc[0]['doc_num']
                comp_grupo = grupo.iloc[0]['comps']
                qtd_total_saida = grupo['qt_entrada'].astype(float).sum()
                
                if doc_grupo in doc_index:
                    candidatos_idx = doc_index[doc_grupo]
                    for idx_e in candidatos_idx:
                        if idx_e in entradas_processadas: continue
                        
                        row_e = df_entrada.loc[idx_e]
                        qtd_e = float(row_e.get('qt_entrada', 0))
                        
                        qtd_match_soma = abs(qtd_e - qtd_total_saida) < 0.1
                        limiar_grupo = 70 if qtd_match_soma else 85
                        
                        score_prod, _ = calcular_similaridade_precalc(comp_grupo, row_e['comps'], ignore_penalties=True)
                        
                        if score_prod >= limiar_grupo:
                            if qtd_match_soma:
                                entradas_processadas.add(idx_e)
                                for idx_s, row_s in grupo.iterrows():
                                    qtd_s = float(row_s.get('qt_entrada', 0))
                                    perc_do_total = qtd_s / qtd_total_saida if qtd_total_saida > 0 else 0
                                    valor_prop_e = float(row_e['valor_total']) * perc_do_total
                                    matches_agrupados[idx_s] = {
                                        'index': idx_e, 'row': row_e, 'score': 100, 'score_produto': score_prod,
                                        'detalhes': f"Agrupado (Soma {len(grupo)} itens)",
                                        'detalhes_produto': "Match por agrupamento de saída",
                                        'valor_entrada_proporcional': valor_prop_e,
                                        'qtd_entrada_proporcional': qtd_s
                                    }
                                break

    stats = {
        'conformes': 0, 'nao_conformes': 0, 'nao_encontrados': 0,
        'valor_divergente': 0, 'qtd_divergente': 0,
        'matches_perfeitos': 0, 'matches_bons': 0, 'matches_razoaveis': 0
    }
    
    total_items = len(df_saida)
    
    for i, (idx_s, row_s) in enumerate(df_saida.iterrows()):
        if progress_callback and i % 20 == 0:
            progress_callback(0.05 + (i / total_items) * 0.9, f"Analisando {i + 1}/{total_items}")

        doc_num = row_s['doc_num']
        produto_s = row_s['ds_produto']
        comp_s = row_s['comps']
        valor_s = float(row_s['valor_total'])
        qtd_s = float(row_s.get('qt_entrada', 0))
        origem_s_norm = row_s['origem_norm']
        destino_s_norm = row_s['destino_norm']
        data_s = row_s['data']
        destino_eh_cp = row_s['destino_cp']
        
        if idx_s in matches_agrupados:
            match_info = matches_agrupados[idx_s]
            row_e = match_info['row']
            valor_e = match_info['valor_entrada_proporcional']
            qtd_e = match_info['qtd_entrada_proporcional']
            diferenca_valor = round(valor_s - valor_e, 2)
            diferenca_qtd = 0
            
            stats['conformes'] += 1
            stats['matches_perfeitos'] += 1
            
            data_e = row_e['data']
            tempo_recebimento = (data_e - data_s).total_seconds() / 3600 if pd.notna(data_s) and pd.notna(data_e) else None
            obs = f"Score:100% | {match_info['detalhes']}"
            if tempo_recebimento is not None and tempo_recebimento < 0:
                obs += " | ⚠️ DATA ANTERIOR (Entrada < Saída)"
            
            analise.append([
                data_s, row_s['unidade_origem'], row_s['unidade_destino'], doc_num, produto_s, row_e['ds_produto'],
                row_s.get('especie', ''), valor_s, valor_e, diferenca_valor, 
                qtd_s, qtd_e, diferenca_qtd,
                data_e, tempo_recebimento, 
                "✅ Conforme", "-", "⭐⭐⭐ Excelente", obs, match_info['detalhes_produto']
            ])
            continue
        
        matches = []
        best_score = 0
        candidatos_idx = []
        match_agregado = None
        documento_nao_encontrado = False
        
        if doc_num and doc_num != '':
            if doc_num in doc_index:
                candidatos_idx = doc_index[doc_num]
                candidatos_disponiveis = [i for i in candidatos_idx if i not in entradas_processadas]
                
                # Agregação One-to-Many logic
                matches_doc_prod = []
                match_exato = None
                
                for idx_e in candidatos_disponiveis:
                    row_e = df_entrada.loc[idx_e]
                    qtd_e = float(row_e.get('qt_entrada', 0))
                    qtd_match_exato = abs(qtd_e - qtd_s) < 0.01
                    limiar_doc = 70 if qtd_match_exato else 85
                    score_prod, _ = calcular_similaridade_precalc(comp_s, row_e['comps'], ignore_penalties=True)
                    if score_prod >= limiar_doc:
                        if qtd_match_exato:
                            match_exato = {
                                'index': idx_e, 'row': row_e, 'score': 100, 'score_produto': score_prod,
                                'detalhes': f"Match exato (Doc:{doc_num})", 'detalhes_produto': "Quantidade exata"
                            }
                        matches_doc_prod.append((idx_e, row_e, score_prod))
                
                if match_exato:
                    match_agregado = match_exato
                elif matches_doc_prod:
                    qtd_total_entrada = sum(float(m[1].get('qt_entrada', 0)) for m in matches_doc_prod)
                    qtd_primeiro = float(matches_doc_prod[0][1].get('qt_entrada', 0))
                    desvio_soma = abs(qtd_total_entrada - qtd_s) / qtd_s * 100 if qtd_s > 0 else 0
                    soma_razoavel = desvio_soma <= 10
                    
                    if soma_razoavel and (len(matches_doc_prod) > 1 or (abs(qtd_total_entrada - qtd_s) < abs(qtd_primeiro - qtd_s))):
                        valor_total_entrada = sum(float(m[1]['valor_total']) for m in matches_doc_prod)
                        row_virtual = matches_doc_prod[0][1].copy()
                        row_virtual['qt_entrada'] = qtd_total_entrada
                        row_virtual['valor_total'] = valor_total_entrada
                        row_virtual['ds_produto'] = f"{row_virtual['ds_produto']} (+ {len(matches_doc_prod)-1} itens)"
                        match_agregado = {
                            'indices': [m[0] for m in matches_doc_prod],
                            'row': row_virtual, 'score': 100, 'score_produto': matches_doc_prod[0][2],
                            'detalhes': f"Agregado: {len(matches_doc_prod)} itens", 'detalhes_produto': "Múltiplos itens somados"
                        }

                candidatos = df_entrada.loc[candidatos_idx]
            else:
                candidatos = pd.DataFrame()
                documento_nao_encontrado = True
        elif destino_eh_cp:
            if pd.notna(data_s):
                mask_data = (df_entrada['data'] >= data_s - pd.Timedelta(days=30)) & \
                            (df_entrada['data'] <= data_s + pd.Timedelta(days=30))
                candidatos = df_entrada[mask_data]
            else:
                candidatos = df_entrada
        else:
            if not documento_nao_encontrado:
                if pd.notna(data_s):
                    mask_data = (df_entrada['data'] >= data_s - pd.Timedelta(days=30)) & \
                                (df_entrada['data'] <= data_s + pd.Timedelta(days=30))
                    candidatos = df_entrada[mask_data]
                else:
                    candidatos = df_entrada
        
        if len(candidatos) > 100: candidatos = candidatos.head(100)
        
        if match_agregado:
            matches.append(match_agregado)
            best_score = 100
        else:
            for idx_e, row_e in candidatos.iterrows():
                if idx_e in entradas_processadas: continue
                if best_score >= 95: break
                
                score_total = 0
                detalhes_match = []
                
                doc_num_e = row_e['doc_num']
                doc_match = False
                
                if destino_eh_cp:
                    score_total += 40
                    detalhes_match.append("Doc:CP")
                    doc_match = True
                elif doc_num and doc_num != '' and doc_num_e and doc_num_e != '':
                    if doc_num == doc_num_e:
                        score_total += 40
                        detalhes_match.append(f"Doc:✓{doc_num}")
                        doc_match = True
                    else:
                        continue
                elif not doc_num or doc_num == '':
                    score_total += 15
                    detalhes_match.append("Doc:N/A(saída)")
                else:
                    score_total += 15
                    detalhes_match.append("Doc:N/A(entrada)")
                
                score_produto, detalhes_produto = calcular_similaridade_precalc(comp_s, row_e['comps'], ignore_penalties=doc_match)
                
                if doc_match:
                    qtd_e = float(row_e.get('qt_entrada', 0))
                    qtd_match_exato = abs(qtd_e - qtd_s) < 0.01
                    limiar_efetivo = 40 if qtd_match_exato else 85
                else:
                    limiar_efetivo = limiar_similaridade
                
                if score_produto < limiar_efetivo: continue
                
                score_total += score_produto * 0.45
                detalhes_match.append(f"Prod:{score_produto:.0f}%")
                
                origem_match = origem_s_norm == row_e['origem_norm']
                destino_match = destino_s_norm == row_e['destino_norm']
                if origem_match or destino_match:
                    score_total += 5
                    detalhes_match.append("Unid:✓")
                
                if pd.notna(data_s) and pd.notna(row_e['data']):
                    diff_dias = abs((row_e['data'] - data_s).days)
                    if diff_dias == 0:
                        score_total += 5
                        detalhes_match.append("Data:mesma")
                    elif diff_dias <= 3:
                        score_total += 4
                
                valor_e = float(row_e['valor_total'])
                if valor_s > 0:
                    perc_diff = abs(valor_s - valor_e) / valor_s * 100
                    if perc_diff <= 1:
                        score_total += 2
                        detalhes_match.append("Valor:≈")
                
                if score_total >= 50:
                    matches.append({
                        'index': idx_e, 'row': row_e, 'score': score_total,
                        'score_produto': score_produto, 'detalhes': ' | '.join(detalhes_match),
                        'detalhes_produto': detalhes_produto
                    })
                    if score_total > best_score: best_score = score_total
        
        if matches:
            matches.sort(key=lambda x: (x['score'], x['score_produto']), reverse=True)
            best_match = matches[0]
            row_e = best_match['row']
            
            valor_e = float(row_e['valor_total'])
            qtd_e = float(row_e.get('qt_entrada', 0))
            
            doc_match = (doc_num and doc_num != '' and doc_num == row_e.get('doc_num', ''))
            is_valid, diferenca_qtd_validada = validar_match_quantidade(qtd_s, qtd_e, best_match['score_produto'], doc_match)
            
            if not is_valid:
                stats['nao_encontrados'] += 1
                stats['nao_conformes'] += 1
                analise.append([
                    data_s, row_s['unidade_origem'], row_s['unidade_destino'], doc_num, produto_s, "-",
                    row_s.get('especie', ''), valor_s, None, None, qtd_s, None, None,
                    None, None, "❌ Não Conforme", "Item não encontrado (Qtd divergente)", "-", 
                    f"Match rejeitado: {qtd_s} vs {qtd_e}", "-"
                ])
                continue
            
            diferenca_qtd = diferenca_qtd_validada
            if 'indices' in best_match:
                for idx in best_match['indices']: entradas_processadas.add(idx)
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
            
            conforme_qtd = abs(diferenca_qtd) < 0.01
            if conforme_qtd:
                if valor_s < 10:
                    conforme_valor = abs(diferenca_valor) <= 1.0
                else:
                    limite_valor_absoluto = max(10.0, valor_s * 0.10)
                    conforme_valor = abs(diferenca_valor) <= limite_valor_absoluto or perc_diff_valor <= 10
            else:
                conforme_valor = abs(diferenca_valor) <= 10
            
            if conforme_valor and conforme_qtd:
                status = "✅ Conforme"
                tipo_div = "-"
                stats['conformes'] += 1
            else:
                status = "❌ Não Conforme"
                stats['nao_conformes'] += 1
                tipos_div = []
                if not conforme_valor:
                    stats['valor_divergente'] += 1
                    tipos_div.append(f"Divergência Valor")
                if not conforme_qtd:
                    stats['qtd_divergente'] += 1
                    tipos_div.append(f"Divergência Qtd")
                tipo_div = " | ".join(tipos_div)
            
            obs = f"Score:{best_match['score']:.0f}% | {best_match['detalhes']}"
            data_e = row_e['data']
            tempo_recebimento = (data_e - data_s).total_seconds() / 3600 if pd.notna(data_s) and pd.notna(data_e) else None
            
            analise.append([
                data_s, row_s['unidade_origem'], row_s['unidade_destino'], doc_num, produto_s, row_e['ds_produto'],
                row_s.get('especie', ''), valor_s, valor_e, diferenca_valor, 
                qtd_s, qtd_e, diferenca_qtd,
                data_e, tempo_recebimento, 
                status, tipo_div, qualidade_match, obs, best_match['detalhes_produto']
            ])
        else:
            stats['nao_encontrados'] += 1
            stats['nao_conformes'] += 1
            motivo = f"Documento {doc_num} não encontrado" if doc_num and not destino_eh_cp else "Item não encontrado"
            analise.append([
                data_s, row_s['unidade_origem'], row_s['unidade_destino'], doc_num, produto_s, "-",
                row_s.get('especie', ''), valor_s, None, None, 
                qtd_s, None, None,
                None, None, 
                "⚠️ Não Recebido", motivo, "-", "Sem correspondência", "-"
            ])
            
    if progress_callback:
        progress_callback(0.95, "Finalizando...")
        
    for idx_e, row_e in df_entrada.iterrows():
        if idx_e in entradas_processadas: continue
        
        data_e = row_e['data']
        if pd.notna(periodo_inicio) and pd.notna(periodo_fim) and pd.notna(data_e):
            if data_e < periodo_inicio or data_e > periodo_fim:
                continue
        
        analise.append([
            row_e['data'], row_e['unidade_origem'], row_e['unidade_destino'],
            row_e['doc_num'], "-", row_e['ds_produto'], row_e.get('especie', ''),
            None, float(row_e['valor_total']), None, 
            None, float(row_e.get('qt_entrada', 0)), None,
            row_e['data'], None, 
            "❌ Não Conforme", "Item recebido sem saída", "-",
            "Entrada órfã", "-"
        ])
        
    df_resultado = pd.DataFrame(analise, columns=[
        "Data", "Unidade Origem", "Unidade Destino", "Documento",
        "Produto (Saída)", "Produto (Entrada)", "Espécie", 
        "Valor Saída (R$)", "Valor Entrada (R$)", "Diferença (R$)",
        "Qtd Saída", "Qtd Entrada", "Diferença Qtd",
        "Data Entrada", "Tempo Recebimento (Horas)",
        "Status", "Tipo de Divergência", 
        "Qualidade Match", "Observações", "Detalhes Produto"
    ])
    
    if progress_callback:
        progress_callback(1.0, "Concluído!")
        
    return df_resultado, stats

def _normalizar_hospital(nome):
    if pd.isna(nome): return nome
    
    # Remove espaços invisíveis/estranhos e normaliza múltiplos espaços
    nome_limpo = " ".join(str(nome).split()).upper()
    
    # Dicionário De/Para conforme especificação do Engenheiro de Dados
    de_para = {
        # HOSPITAL CASA DE PORTUGAL
        'CASA DE PORTUGAL': 'HOSPITAL CASA DE PORTUGAL',
        'CASA DE PORTUGAL - REDE CASA': 'HOSPITAL CASA DE PORTUGAL',
        
        # HOSPITAL CASA MENSSANA
        'HOSPITAL CASA MENSSANA - REDE CASA': 'HOSPITAL CASA MENSSANA',
        'HC MENSSANA PARTICULAR - REDE CASA': 'HOSPITAL CASA MENSSANA',
        
        # HOSPITAL CASA EVANGELICO
        'HOSPITAL EVANGELICO - REDE CASA': 'HOSPITAL CASA EVANGELICO',
        'HOSPITAL CASA EVANGÉLICO - REDE CASA': 'HOSPITAL CASA EVANGELICO',
        'HOSP.EVANGELICO - REDE CASA': 'HOSPITAL CASA EVANGELICO',
        'HOSPITAL CASA EVANGELICO - REDE CASA': 'HOSPITAL CASA EVANGELICO',
        
        # HOSPITAL CASA RIO LARANJEIRAS
        'HOSPITAL CASA RIO LARANJEIRAS - REDE CASA': 'HOSPITAL CASA RIO LARANJEIRAS',
        'HOSPITAL RIO LARANJEIRAS - REDE CASA': 'HOSPITAL CASA RIO LARANJEIRAS',
        'HOSPITAL RIO LARANJEIRAS LTDA - REDE CASA': 'HOSPITAL CASA RIO LARANJEIRAS',
        
        # HOSPITAL CASA RIO BOTAFOGO
        'HOSPITAL CASA RIO BOTAFOGO - REDE CASA': 'HOSPITAL CASA RIO BOTAFOGO',
        
        # HOSPITAL CASA SANTA CRUZ
        'HOSPITAL CASA SANTA CRUZ - REDE CASA': 'HOSPITAL CASA SANTA CRUZ',
        'HOSPITAL SANTA CRUZ - REDE CASA': 'HOSPITAL CASA SANTA CRUZ',
        'HOSPITAL SANTA CRUZ': 'HOSPITAL CASA SANTA CRUZ',
        
        # HOSPITAL CASA SAO BERNARDO
        'HOSPITAL CASA SAO BERNARDO - REDE CASA': 'HOSPITAL CASA SAO BERNARDO',
        
        # HOSPITAL CASA PREMIUM
        'HOSPITAL DE CANCER': 'HOSPITAL CASA PREMIUM',
        'HOSPITAL DE CANCER - REDE CASA': 'HOSPITAL CASA PREMIUM',
        'HOSPITAL CASA HOSPITAL DO CANCER – HCHC ADMINISTRACAO E GEST - REDE CASA': 'HOSPITAL CASA PREMIUM',
        'HOSPITAL CASA HOSPITAL DO CANCER - REDE CASA': 'HOSPITAL CASA PREMIUM',
        
        # HOSPITAL CASA ILHA DO GOVERNADOR
        'HOSPITAL ILHA DO GOVERNADOR': 'HOSPITAL CASA ILHA DO GOVERNADOR',
        'HOSPITAL ILHA DO GOVERNADOR - REDE CASA': 'HOSPITAL CASA ILHA DO GOVERNADOR',
        'HOSPITAL ILHA DO GOVERNADOR LTDA - REDE CASA': 'HOSPITAL CASA ILHA DO GOVERNADOR'
    }
    
    # Retorna o valor mapeado ou o original se não encontrar
    return de_para.get(nome_limpo, nome_limpo)

def _combinar_data_hora(df):
    if 'hora' in df.columns and 'data' in df.columns:
        df['data'] = pd.to_datetime(df['data'], errors='coerce')
        
        if not pd.api.types.is_timedelta64_dtype(df['hora']):
            mask_time = df['hora'].apply(lambda x: hasattr(x, 'hour'))
            if mask_time.any():
                df.loc[mask_time, 'hora'] = df.loc[mask_time, 'hora'].apply(
                    lambda x: f"{x.hour:02d}:{x.minute:02d}:{x.second:02d}"
                )
            df['hora'] = pd.to_timedelta(df['hora'].astype(str), errors='coerce').fillna(pd.Timedelta(0))
        
        df['data'] = df['data'] + df['hora']
        
        try:
            df['data'] = df['data'].dt.tz_localize('America/Sao_Paulo', ambiguous='NaT', nonexistent='shift_forward')
        except Exception:
            df['data'] = df['data'].dt.tz_convert('America/Sao_Paulo')
            
    return df

def mapear_colunas(df):
    mapeamento = {}
    for col in df.columns:
        col_lower = col.lower()
        if any(x in col_lower for x in ['produto', 'descrição', 'descricao', 'material']):
            if 'ds_produto' not in mapeamento.values(): mapeamento[col] = 'ds_produto'
        elif any(x in col_lower for x in ['documento', 'nf', 'nota']):
            if 'documento' not in mapeamento.values(): mapeamento[col] = 'documento'
        elif 'origem' in col_lower and 'unidade' in col_lower:
            mapeamento[col] = 'unidade_origem'
        elif 'destino' in col_lower and 'unidade' in col_lower:
            mapeamento[col] = 'unidade_destino'
        elif any(x in col_lower for x in ['valor total', 'vl_total', 'total']):
            if 'valor_total' not in mapeamento.values(): mapeamento[col] = 'valor_total'
        elif any(x in col_lower for x in ['quantidade', 'qtd', 'qt_entrada', 'qt entrada']):
            if 'qt_entrada' not in mapeamento.values(): mapeamento[col] = 'qt_entrada'
        elif 'especie' in col_lower or 'espécie' in col_lower:
            if 'especie' not in mapeamento.values(): mapeamento[col] = 'especie'
        elif any(x in col_lower for x in ['hora', 'time']):
            if 'hora' not in mapeamento.values(): mapeamento[col] = 'hora'
    return mapeamento

def preparar_dataframe(df):
    df.columns = [c.strip().lower() for c in df.columns]
    map_cols = mapear_colunas(df)
    df.rename(columns=map_cols, inplace=True)
    df = _combinar_data_hora(df)
    df['data'] = _parse_date_column(df['data'])
    
    if 'unidade_origem' in df.columns:
        df['unidade_origem'] = df['unidade_origem'].apply(_normalizar_hospital)
    if 'unidade_destino' in df.columns:
        df['unidade_destino'] = df['unidade_destino'].apply(_normalizar_hospital)
        
    # Normalização numérica
    for col in ['qt_entrada', 'valor_total']:
        if col in df.columns:
            df[col] = df[col].apply(normalizar_valor_numerico)
            
    # Filtro Oftalmocasa
    termo_exclusao = "OFTALMOCASA"
    if 'unidade_origem' in df.columns and 'unidade_destino' in df.columns:
        mask = df['unidade_origem'].str.contains(termo_exclusao, na=False) | \
               df['unidade_destino'].str.contains(termo_exclusao, na=False)
        df = df[~mask]
        
    return df
