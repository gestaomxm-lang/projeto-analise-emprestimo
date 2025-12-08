import pandas as pd
import os
import re
from difflib import SequenceMatcher
from datetime import timedelta

def extrair_numeros(documento):
    """Extrai a primeira sequência numérica encontrada no documento."""
    match = re.search(r'\d+', str(documento))
    return match.group() if match else ""

def _parse_date_column(series):
    """Converte a coluna de datas para datetime."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_datetime(series, unit='d', origin='1899-12-30', errors='coerce')
    return pd.to_datetime(series, dayfirst=True, errors='coerce')

def extrair_componentes_produto(descricao):
    """
    Extrai componentes principais do produto para matching inteligente.
    Retorna: princípio ativo, concentração, apresentação, fabricante
    """
    descricao = str(descricao).upper().strip()
    
    componentes = {
        'original': descricao,
        'normalizado': '',
        'principio_ativo': '',
        'concentracao': '',
        'apresentacao': '',
        'quantidade': '',
        'unidade_medida': '',
        'palavras_chave': []
    }
    
    # Remove pontuações e normaliza
    texto_limpo = re.sub(r'[^\w\s]', ' ', descricao)
    texto_limpo = re.sub(r'\s+', ' ', texto_limpo).strip()
    componentes['normalizado'] = texto_limpo
    
    # Extrai concentração (ex: 500MG, 10ML, 2,5G, etc)
    concentracoes = re.findall(r'\d+[,.]?\d*\s*(?:MG|G|ML|MCG|UI|%|MG/ML|G/ML)', descricao)
    if concentracoes:
        componentes['concentracao'] = ' '.join(concentracoes)
    
    # Extrai apresentação (ampola, comprimido, frasco, etc)
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
    
    # Extrai palavras-chave (remove stopwords médicas comuns)
    stopwords = [
        'DE', 'DA', 'DO', 'COM', 'PARA', 'EM', 'A', 'O', 'E', 'C/',
        'SOLUCAO', 'SOL', 'INJETAVEL', 'INJ', 'ORAL', 'USO', 'ADULTO',
        'PEDIATRICO', 'ESTERIL', 'DESCARTAVEL', 'DESC'
    ]
    palavras = texto_limpo.split()
    palavras_chave = [p for p in palavras if p not in stopwords and len(p) > 2]
    componentes['palavras_chave'] = palavras_chave[:5]  # Primeiras 5 palavras relevantes
    
    # Tenta identificar princípio ativo (geralmente primeiras palavras)
    if len(palavras_chave) > 0:
        componentes['principio_ativo'] = ' '.join(palavras_chave[:2])
    
    return componentes

def calcular_similaridade_produtos(prod1, prod2):
    """
    Calcula similaridade entre dois produtos considerando múltiplos fatores.
    Retorna score de 0 a 100 e detalhes do matching.
    """
    comp1 = extrair_componentes_produto(prod1)
    comp2 = extrair_componentes_produto(prod2)
    
    score = 0
    detalhes = []
    
    # 1. Similaridade textual geral (peso: 30%)
    sim_geral = SequenceMatcher(None, comp1['normalizado'], comp2['normalizado']).ratio()
    score += sim_geral * 30
    detalhes.append(f"Texto:{sim_geral:.0%}")
    
    # 2. Princípio ativo (peso: 35%)
    if comp1['principio_ativo'] and comp2['principio_ativo']:
        sim_principio = SequenceMatcher(None, comp1['principio_ativo'], comp2['principio_ativo']).ratio()
        score += sim_principio * 35
        detalhes.append(f"Princípio:{sim_principio:.0%}")
    
    # 3. Concentração (peso: 20%)
    if comp1['concentracao'] and comp2['concentracao']:
        if comp1['concentracao'] == comp2['concentracao']:
            score += 20
            detalhes.append(f"Conc:✓")
        else:
            # Verifica se concentrações são similares
            sim_conc = SequenceMatcher(None, comp1['concentracao'], comp2['concentracao']).ratio()
            if sim_conc > 0.7:
                score += sim_conc * 20
                detalhes.append(f"Conc:~{sim_conc:.0%}")
    
    # 4. Apresentação (peso: 10%)
    if comp1['apresentacao'] and comp2['apresentacao']:
        if comp1['apresentacao'] == comp2['apresentacao']:
            score += 10
            detalhes.append(f"Apres:✓")
        else:
            # Verifica apresentações equivalentes
            equiv_apresentacao = {
                'AMPOLA': ['AMP', 'AMPOLA'],
                'COMPRIMIDO': ['COMP', 'CP', 'COMPRIMIDO'],
                'CAPSULA': ['CAPS', 'CAPSULA'],
                'FRASCO': ['FR', 'FRASCO'],
                'SERINGA': ['SER', 'SERINGA']
            }
            for grupo in equiv_apresentacao.values():
                if comp1['apresentacao'] in grupo and comp2['apresentacao'] in grupo:
                    score += 10
                    detalhes.append(f"Apres:equiv")
                    break
    
    # 5. Palavras-chave em comum (peso: 5%)
    palavras_comum = set(comp1['palavras_chave']) & set(comp2['palavras_chave'])
    if palavras_comum:
        perc_comum = len(palavras_comum) / max(len(comp1['palavras_chave']), len(comp2['palavras_chave']))
        score += perc_comum * 5
        detalhes.append(f"Palavras:{len(palavras_comum)}")
    
    return score, ' | '.join(detalhes), comp1, comp2

def eh_casa_portugal(unidade):
    """Verifica se a unidade é Casa de Portugal."""
    unidade_norm = str(unidade).upper().strip()
    return 'CASA' in unidade_norm and 'PORTUGAL' in unidade_norm

def carregar_dados():
    """Carrega os arquivos da pasta 'projeto análise de empréstimo' no Desktop."""
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    projeto_path = os.path.join(desktop_path, "projeto análise de empréstimo")

    saida_path = os.path.join(projeto_path, "emprestimo_concedido.xlsx")
    entrada_path = os.path.join(projeto_path, "emprestimo_recebido.xlsx")

    if not os.path.exists(saida_path) or not os.path.exists(entrada_path):
        print("❌ Erro: arquivos não encontrados na pasta 'projeto análise de empréstimo'.")
        print(f"   Procurando em: {projeto_path}")
        print(f"   Arquivo saída: {saida_path} - {'Existe' if os.path.exists(saida_path) else 'NÃO EXISTE'}")
        print(f"   Arquivo entrada: {entrada_path} - {'Existe' if os.path.exists(entrada_path) else 'NÃO EXISTE'}")
        return None, None

    try:
        df_saida = pd.read_excel(saida_path, engine='openpyxl')
        df_entrada = pd.read_excel(entrada_path, engine='openpyxl')
    except Exception as e:
        print(f"❌ Erro ao ler arquivos Excel: {e}")
        return None, None

    # Padroniza nomes das colunas
    df_saida.columns = [c.strip().lower() for c in df_saida.columns]
    df_entrada.columns = [c.strip().lower() for c in df_entrada.columns]

    # Converte datas
    df_saida['data'] = _parse_date_column(df_saida['data'])
    df_entrada['data'] = _parse_date_column(df_entrada['data'])

    return df_saida, df_entrada

def buscar_correspondencia(row_saida, df_entrada, limiar_similaridade=65):
    """
    Busca correspondência para um item de saída na entrada.
    Usa matching inteligente considerando equivalências de descrição.
    """
    doc_num = extrair_numeros(row_saida['documento'])
    produto_s = row_saida['ds_produto']
    valor_s = float(row_saida['valor_total'])
    origem_s = row_saida['unidade_origem']
    destino_s = row_saida['unidade_destino']
    data_s = row_saida['data']
    
    # Verifica se destino é Casa de Portugal
    destino_eh_cp = eh_casa_portugal(destino_s)
    
    matches = []
    
    for idx, row_e in df_entrada.iterrows():
        doc_num_e = extrair_numeros(row_e['documento'])
        produto_e = row_e['ds_produto']
        valor_e = float(row_e['valor_total'])
        origem_e = row_e['unidade_origem']
        destino_e = row_e['unidade_destino']
        data_e = row_e['data']
        
        # Calcula score de correspondência
        score_total = 0
        detalhes_match = []
        
        # 1. Similaridade de produto (peso: 50%)
        score_produto, detalhes_produto, comp_s, comp_e = calcular_similaridade_produtos(produto_s, produto_e)
        
        if score_produto < limiar_similaridade:
            continue  # Se produto não é similar o suficiente, pula
        
        score_total += score_produto * 0.5
        detalhes_match.append(f"Prod:{score_produto:.0f}%")
        
        # 2. Correspondência de documento (peso: 25%)
        if destino_eh_cp:
            # Casa de Portugal: não exige documento
            score_total += 25
            detalhes_match.append("Doc:CP")
        elif doc_num and doc_num_e:
            if doc_num == doc_num_e:
                score_total += 25
                detalhes_match.append(f"Doc:{doc_num}")
            else:
                # Documentos diferentes - penaliza mas não elimina
                score_total += 5
                detalhes_match.append(f"Doc:diff")
        else:
            # Sem documento - considera outros fatores
            score_total += 10
            detalhes_match.append("Doc:N/A")
        
        # 3. Correspondência de unidades (peso: 10%)
        origem_match = origem_s.upper().strip() == origem_e.upper().strip()
        destino_match = destino_s.upper().strip() == destino_e.upper().strip()
        
        if origem_match and destino_match:
            score_total += 10
            detalhes_match.append("Unid:✓✓")
        elif origem_match or destino_match:
            score_total += 5
            detalhes_match.append("Unid:~")
        
        # 4. Proximidade de data (peso: 10%)
        if pd.notna(data_s) and pd.notna(data_e):
            diff_dias = abs((data_e - data_s).days)
            if diff_dias == 0:
                score_total += 10
                detalhes_match.append("Data:mesma")
            elif diff_dias <= 3:
                score_total += 8
                detalhes_match.append(f"Data:{diff_dias}d")
            elif diff_dias <= 7:
                score_total += 5
                detalhes_match.append(f"Data:{diff_dias}d")
            elif diff_dias <= 15:
                score_total += 2
                detalhes_match.append(f"Data:{diff_dias}d")
        
        # 5. Proximidade de valor (peso: 5%)
        if valor_s > 0:
            diff_valor = abs(valor_s - valor_e)
            perc_diff = (diff_valor / valor_s * 100)
            
            if perc_diff <= 1:
                score_total += 5
                detalhes_match.append("Valor:≈")
            elif perc_diff <= 5:
                score_total += 4
                detalhes_match.append(f"Valor:~{perc_diff:.1f}%")
            elif perc_diff <= 15:
                score_total += 2
                detalhes_match.append(f"Valor:~{perc_diff:.1f}%")
            elif perc_diff <= 50:
                score_total += 1
                detalhes_match.append(f"Valor:diff{perc_diff:.0f}%")
        
        # Score mínimo para considerar match
        if score_total >= 50:
            matches.append({
                'index': idx,
                'row': row_e,
                'score': score_total,
                'score_produto': score_produto,
                'detalhes': ' | '.join(detalhes_match),
                'detalhes_produto': detalhes_produto,
                'comp_saida': comp_s,
                'comp_entrada': comp_e
            })
    
    # Ordena por score decrescente
    matches.sort(key=lambda x: (x['score'], x['score_produto']), reverse=True)
    return matches

def analisar_itens(df_saida, df_entrada, limiar_similaridade=65):
    """Compara item a item entre saída e entrada com matching inteligente."""
    analise = []
    entradas_processadas = set()
    
    # Normaliza colunas e remove espaços extras
    for df in [df_saida, df_entrada]:
        df['documento'] = df['documento'].astype(str).str.strip()
        df['ds_produto'] = df['ds_produto'].astype(str).str.strip()
        df['unidade_origem'] = df['unidade_origem'].astype(str).str.strip()
        df['unidade_destino'] = df['unidade_destino'].astype(str).str.strip()
    
    print(f"\n🔍 Analisando {len(df_saida)} itens de saída...")
    print(f"⚙️  Limiar de similaridade: {limiar_similaridade}%")
    
    # Contadores para estatísticas
    stats = {
        'conformes': 0,
        'nao_conformes': 0,
        'nao_encontrados': 0,
        'valor_divergente': 0,
        'matches_perfeitos': 0,
        'matches_bons': 0,
        'matches_razoaveis': 0
    }
    
    # Loop por cada item de saída
    for idx_s, row_s in df_saida.iterrows():
        doc_num = extrair_numeros(row_s['documento'])
        produto_s = row_s['ds_produto']
        valor_s = float(row_s['valor_total'])
        especie = row_s.get('especie', '')
        origem_s = row_s['unidade_origem']
        destino_s = row_s['unidade_destino']
        data_s = row_s['data']
        
        # Busca correspondências
        matches = buscar_correspondencia(row_s, df_entrada, limiar_similaridade)
        
        if matches:
            # Pega o melhor match
            best_match = matches[0]
            row_e = best_match['row']
            idx_e = best_match['index']
            
            entradas_processadas.add(idx_e)
            
            valor_e = float(row_e['valor_total'])
            diferenca = round(valor_s - valor_e, 2)
            perc_diff = abs(diferenca / valor_s * 100) if valor_s > 0 else 0
            
            # Classifica qualidade do match
            if best_match['score'] >= 90:
                qualidade_match = "⭐⭐⭐ Excelente"
                stats['matches_perfeitos'] += 1
            elif best_match['score'] >= 75:
                qualidade_match = "⭐⭐ Bom"
                stats['matches_bons'] += 1
            else:
                qualidade_match = "⭐ Razoável"
                stats['matches_razoaveis'] += 1
            
            # Determina conformidade
            conforme = abs(diferenca) <= 10
            
            if conforme:
                status = "✅ Conforme"
                tipo_div = "-"
                stats['conformes'] += 1
            else:
                status = "⚠️ Não Conforme"
                stats['nao_conformes'] += 1
                stats['valor_divergente'] += 1
                if perc_diff > 50:
                    tipo_div = f"Valor muito divergente ({perc_diff:.1f}%)"
                elif perc_diff > 20:
                    tipo_div = f"Valor divergente ({perc_diff:.1f}%)"
                else:
                    tipo_div = f"Pequena divergência ({perc_diff:.1f}%)"
            
            # Monta observações detalhadas
            obs_partes = [
                f"Score:{best_match['score']:.0f}%",
                qualidade_match,
                best_match['detalhes']
            ]
            
            if len(matches) > 1:
                obs_partes.append(f"({len(matches)} possíveis)")
            
            obs = " | ".join(obs_partes)
            
            # Informações sobre componentes do produto
            comp_info = f"{best_match['detalhes_produto']}"
            
            analise.append([
                data_s,
                origem_s,
                destino_s,
                doc_num,
                produto_s,
                row_e['ds_produto'],
                especie,
                valor_s,
                valor_e,
                diferenca,
                status,
                tipo_div,
                qualidade_match,
                obs,
                comp_info
            ])
        else:
            # Não encontrou correspondência
            stats['nao_encontrados'] += 1
            stats['nao_conformes'] += 1
            
            analise.append([
                data_s,
                origem_s,
                destino_s,
                doc_num,
                produto_s,
                "-",
                especie,
                valor_s,
                None,
                None,
                "❌ Não Conforme",
                "Item não encontrado na entrada",
                "-",
                "Sem correspondência encontrada",
                "-"
            ])
        
        # Progresso
        if (idx_s + 1) % 50 == 0:
            print(f"   Processados: {idx_s + 1}/{len(df_saida)}")
    
    print(f"✅ Análise de saídas concluída!")
    print(f"\n🔍 Verificando entradas sem saída correspondente...")
    
    # Procura itens recebidos sem saída correspondente
    entradas_orfas = 0
    for idx_e, row_e in df_entrada.iterrows():
        if idx_e in entradas_processadas:
            continue
        
        entradas_orfas += 1
        doc_num_e = extrair_numeros(row_e['documento'])
        produto_e = row_e['ds_produto']
        
        analise.append([
            row_e['data'],
            row_e['unidade_origem'],
            row_e['unidade_destino'],
            doc_num_e,
            "-",
            produto_e,
            row_e.get('especie', ''),
            None,
            float(row_e['valor_total']),
            None,
            "❌ Não Conforme",
            "Item recebido sem saída correspondente",
            "-",
            "Entrada órfã - possível erro de lançamento",
            "-"
        ])
    
    print(f"   Encontradas {entradas_orfas} entradas sem saída correspondente")
    
    df_resultado = pd.DataFrame(analise, columns=[
        "Data", "Unidade Origem", "Unidade Destino", "Documento",
        "Produto (Saída)", "Produto (Entrada)", "Espécie", 
        "Valor Saída (R$)", "Valor Entrada (R$)",
        "Diferença (R$)", "Status", "Tipo de Divergência", 
        "Qualidade Match", "Observações", "Detalhes Produto"
    ])
    
    return df_resultado, stats

def gerar_estatisticas(df, stats):
    """Gera estatísticas detalhadas da análise."""
    total = len(df)
    
    print(f"\n{'='*70}")
    print(f"📊 ESTATÍSTICAS DA ANÁLISE")
    print(f"{'='*70}")
    print(f"\n📈 RESUMO GERAL:")
    print(f"   Total de registros analisados: {total}")
    print(f"   ✅ Conformes: {stats['conformes']} ({stats['conformes']/total*100:.1f}%)")
    print(f"   ❌ Não conformes: {stats['nao_conformes']} ({stats['nao_conformes']/total*100:.1f}%)")
    
    print(f"\n🎯 QUALIDADE DOS MATCHES:")
    print(f"   ⭐⭐⭐ Matches excelentes (>90%): {stats['matches_perfeitos']}")
    print(f"   ⭐⭐ Matches bons (75-90%): {stats['matches_bons']}")
    print(f"   ⭐ Matches razoáveis (50-75%): {stats['matches_razoaveis']}")
    print(f"   ❌ Não encontrados: {stats['nao_encontrados']}")
    
    if stats['nao_conformes'] > 0:
        print(f"\n⚠️  TIPOS DE DIVERGÊNCIA:")
        divergencias = df[df['Status'].str.contains('Não Conforme', na=False)]['Tipo de Divergência'].value_counts()
        for tipo, qtd in divergencias.head(10).items():
            print(f"      • {tipo}: {qtd}")
    
    print(f"\n{'='*70}")

def salvar_relatorio(df):
    """Salva o relatório Excel formatado com múltiplas abas."""
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    projeto_path = os.path.join(desktop_path, "projeto análise de empréstimo")
    output_path = os.path.join(projeto_path, "analise_emprestimos_detalhada.xlsx")

    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            # Aba principal com todos os dados
            df.to_excel(writer, index=False, sheet_name="Análise Completa")
            
            # Aba só com não conformes
            df_nao_conforme = df[df['Status'].str.contains('Não Conforme', na=False)]
            df_nao_conforme.to_excel(writer, index=False, sheet_name="Não Conformes")
            
            # Aba só com conformes
            df_conforme = df[df['Status'].str.contains('Conforme', na=False) & ~df['Status'].str.contains('Não', na=False)]
            df_conforme.to_excel(writer, index=False, sheet_name="Conformes")

            workbook = writer.book
            
            # Formatos
            money_fmt = workbook.add_format({'num_format': 'R$ #,##0.00'})
            date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy'})
            conforme_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'bold': True})
            nao_conforme_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'bold': True})
            header_fmt = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white'})

            # Formata cada aba
            for sheet_name in ["Análise Completa", "Não Conformes", "Conformes"]:
                if sheet_name in writer.sheets:
                    ws = writer.sheets[sheet_name]
                    
                    # Larguras das colunas
                    ws.set_column('A:A', 12, date_fmt)
                    ws.set_column('B:D', 25)
                    ws.set_column('E:F', 55)  # Produtos
                    ws.set_column('G:G', 20)  # Espécie
                    ws.set_column('H:J', 18, money_fmt)
                    ws.set_column('K:K', 18)  # Status
                    ws.set_column('L:L', 35)  # Tipo divergência
                    ws.set_column('M:M', 20)  # Qualidade match
                    ws.set_column('N:N', 50)  # Observações
                    ws.set_column('O:O', 40)  # Detalhes produto
                    
                    # Congela primeira linha
                    ws.freeze_panes(1, 0)
                    
                    # Formata cabeçalho
                    for col_num, value in enumerate(df.columns.values):
                        ws.write(0, col_num, value, header_fmt)
                    
                    # Formatação condicional para status
                    df_sheet = df if sheet_name == "Análise Completa" else (df_nao_conforme if sheet_name == "Não Conformes" else df_conforme)
                    for row_num in range(len(df_sheet)):
                        status = df_sheet.iloc[row_num]['Status']
                        if 'Conforme' in status and 'Não' not in status:
                            ws.write(row_num + 1, 10, status, conforme_fmt)
                        else:
                            ws.write(row_num + 1, 10, status, nao_conforme_fmt)

        print(f"\n✅ Relatório salvo com sucesso!")
        print(f"📁 Local: {output_path}")
        print(f"📊 Abas criadas:")
        print(f"   • Análise Completa: Todos os registros")
        print(f"   • Não Conformes: Apenas divergências")
        print(f"   • Conformes: Apenas itens conformes")
        
    except PermissionError:
        print(f"\n❌ Erro: Não foi possível salvar o arquivo.")
        print(f"⚠️  O arquivo '{output_path}' pode estar aberto.")
        print(f"💡 Solução: Feche o arquivo Excel e execute novamente.")
        raise
    except Exception as e:
        print(f"\n❌ Erro ao salvar o relatório: {e}")
        raise

def executar():
    print("\n" + "="*70)
    print("🏥 ANÁLISE INTELIGENTE DE EMPRÉSTIMOS ENTRE HOSPITAIS")
    print("="*70)
    print("✨ Com matching de equivalências de produtos e apresentações")
    print("="*70)
    
    try:
        df_saida, df_entrada = carregar_dados()
        if df_saida is None or df_entrada is None:
            return

        # Limiar de similaridade (65% = aceita produtos com descrições diferentes mas equivalentes)
        limiar = 65
        
        df_resultado, stats = analisar_itens(df_saida, df_entrada, limiar)
        gerar_estatisticas(df_resultado, stats)
        salvar_relatorio(df_resultado)
        
        print("\n" + "="*70)
        print("✅ PROCESSO CONCLUÍDO COM SUCESSO!")
        print("="*70 + "\n")
        
    except Exception as e:
        # Em caso de erro, salvar em arquivo de log
        log_path = os.path.join(os.path.expanduser("~"), "Desktop", "projeto análise de empréstimo", "erro_log.txt")
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                import traceback
                f.write(f"Erro durante a execução: {e}\n\n")
                f.write(traceback.format_exc())
            print(f"\n❌ Erro durante execução. Log salvo em: {log_path}")
        except:
            print(f"\n❌ Erro durante execução: {e}")
        raise

if __name__ == "__main__":
    executar()