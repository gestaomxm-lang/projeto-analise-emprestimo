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

def normalizar_descricao(texto):
    """Normaliza a descrição do produto para comparação."""
    texto = str(texto).lower().strip()
    # Remove pontuações e caracteres especiais
    texto = re.sub(r'[^\w\s]', ' ', texto)
    # Remove espaços múltiplos
    texto = re.sub(r'\s+', ' ', texto)
    return texto

def calcular_similaridade(texto1, texto2):
    """Calcula a similaridade entre duas strings (0 a 1)."""
    return SequenceMatcher(None, normalizar_descricao(texto1), normalizar_descricao(texto2)).ratio()

def eh_casa_portugal(unidade):
    """Verifica se a unidade é Casa de Portugal."""
    unidade_norm = str(unidade).lower().strip()
    return 'casa' in unidade_norm and 'portugal' in unidade_norm

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

def buscar_correspondencia(row_saida, df_entrada, limiar_similaridade=0.75):
    """
    Busca correspondência para um item de saída na entrada.
    Retorna lista de matches ordenados por score de confiança.
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
        score = 0
        detalhes = []
        
        # 1. Similaridade de descrição (peso: 40%)
        sim_descricao = calcular_similaridade(produto_s, produto_e)
        if sim_descricao >= limiar_similaridade:
            score += sim_descricao * 40
            detalhes.append(f"Desc:{sim_descricao:.2f}")
        else:
            continue  # Se descrição não é similar, pula
        
        # 2. Correspondência de documento (peso: 30%)
        # Para Casa de Portugal, não considera documento
        if destino_eh_cp:
            score += 30  # Dá pontuação máxima se for Casa de Portugal
            detalhes.append("CP:sem_doc")
        elif doc_num and doc_num_e and doc_num == doc_num_e:
            score += 30
            detalhes.append(f"Doc:{doc_num}")
        elif doc_num and doc_num_e:
            continue  # Se tem documento mas não bate, pula
        
        # 3. Correspondência de unidades (peso: 15%)
        if normalizar_descricao(origem_s) == normalizar_descricao(origem_e):
            score += 7.5
            detalhes.append("Origem:OK")
        if normalizar_descricao(destino_s) == normalizar_descricao(destino_e):
            score += 7.5
            detalhes.append("Destino:OK")
        
        # 4. Proximidade de data (peso: 10%)
        if pd.notna(data_s) and pd.notna(data_e):
            diff_dias = abs((data_e - data_s).days)
            if diff_dias <= 7:  # Até 7 dias de diferença
                score += 10 * (1 - diff_dias/7)
                detalhes.append(f"Data:{diff_dias}d")
        
        # 5. Proximidade de valor (peso: 5%)
        diff_valor = abs(valor_s - valor_e)
        perc_diff = (diff_valor / valor_s * 100) if valor_s > 0 else 100
        if perc_diff <= 10:  # Até 10% de diferença
            score += 5 * (1 - perc_diff/10)
            detalhes.append(f"Valor:{perc_diff:.1f}%")
        
        if score >= 50:  # Score mínimo para considerar match
            matches.append({
                'index': idx,
                'row': row_e,
                'score': score,
                'detalhes': ' | '.join(detalhes),
                'similaridade_desc': sim_descricao
            })
    
    # Ordena por score decrescente
    matches.sort(key=lambda x: x['score'], reverse=True)
    return matches

def analisar_itens(df_saida, df_entrada, limiar_similaridade=0.75):
    """Compara item a item entre saída e entrada com matching fuzzy."""
    analise = []
    entradas_processadas = set()
    
    # Normaliza colunas e remove espaços extras
    for df in [df_saida, df_entrada]:
        df['documento'] = df['documento'].astype(str).str.strip()
        df['ds_produto'] = df['ds_produto'].astype(str).str.strip()
        df['unidade_origem'] = df['unidade_origem'].astype(str).str.strip()
        df['unidade_destino'] = df['unidade_destino'].astype(str).str.strip()
    
    print(f"\n🔍 Analisando {len(df_saida)} itens de saída...")
    
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
            
            # Determina conformidade
            conforme = abs(diferenca) <= 10
            
            if conforme:
                status = "✅ Conforme"
                tipo_div = "-"
            else:
                status = "⚠️ Não Conforme"
                if perc_diff > 50:
                    tipo_div = f"Valor muito divergente ({perc_diff:.1f}%)"
                else:
                    tipo_div = f"Valor divergente ({perc_diff:.1f}%)"
            
            # Adiciona informações de matching
            obs = f"Match: {best_match['score']:.0f}% | Sim: {best_match['similaridade_desc']:.0%}"
            if len(matches) > 1:
                obs += f" | {len(matches)} possíveis"
            
            analise.append([
                data_s,
                origem_s,
                destino_s,
                doc_num,
                produto_s,
                row_e['ds_produto'],  # Descrição na entrada
                especie,
                valor_s,
                valor_e,
                diferenca,
                status,
                tipo_div,
                obs
            ])
        else:
            # Não encontrou correspondência
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
                "Sem correspondência"
            ])
    
    print(f"✅ Análise de saídas concluída!")
    print(f"\n🔍 Verificando entradas sem saída correspondente...")
    
    # Procura itens recebidos sem saída correspondente
    for idx_e, row_e in df_entrada.iterrows():
        if idx_e in entradas_processadas:
            continue
        
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
            "Entrada órfã"
        ])
    
    df_resultado = pd.DataFrame(analise, columns=[
        "Data", "Unidade Origem", "Unidade Destino", "Documento",
        "Produto (Saída)", "Produto (Entrada)", "Espécie", 
        "Valor Saída (R$)", "Valor Entrada (R$)",
        "Diferença (R$)", "Status", "Tipo de Divergência", "Observações"
    ])
    
    return df_resultado

def gerar_estatisticas(df):
    """Gera estatísticas da análise."""
    total = len(df)
    conformes = len(df[df['Status'].str.contains('Conforme', na=False)])
    nao_conformes = total - conformes
    
    print(f"\n📊 ESTATÍSTICAS DA ANÁLISE:")
    print(f"   Total de registros: {total}")
    print(f"   ✅ Conformes: {conformes} ({conformes/total*100:.1f}%)")
    print(f"   ❌ Não conformes: {nao_conformes} ({nao_conformes/total*100:.1f}%)")
    
    if nao_conformes > 0:
        print(f"\n   Tipos de divergência:")
        divergencias = df[df['Status'].str.contains('Não Conforme', na=False)]['Tipo de Divergência'].value_counts()
        for tipo, qtd in divergencias.items():
            print(f"      • {tipo}: {qtd}")

def salvar_relatorio(df):
    """Salva o relatório Excel formatado."""
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    projeto_path = os.path.join(desktop_path, "projeto análise de empréstimo")
    output_path = os.path.join(projeto_path, "analise_emprestimos_detalhada.xlsx")

    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name="Análise Detalhada")

            workbook = writer.book
            ws = writer.sheets["Análise Detalhada"]

            # Formatos
            money_fmt = workbook.add_format({'num_format': 'R$ #,##0.00'})
            date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy'})
            conforme_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
            nao_conforme_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

            # Larguras das colunas
            ws.set_column('A:A', 12, date_fmt)
            ws.set_column('B:D', 25)
            ws.set_column('E:F', 50)  # Produtos
            ws.set_column('G:G', 25)
            ws.set_column('H:J', 18, money_fmt)
            ws.set_column('K:K', 20)
            ws.set_column('L:L', 35)
            ws.set_column('M:M', 40)

            # Formatação condicional para status
            for row_num in range(1, len(df) + 1):
                status = df.iloc[row_num - 1]['Status']
                if 'Conforme' in status and 'Não' not in status:
                    ws.write(row_num, 10, status, conforme_fmt)
                else:
                    ws.write(row_num, 10, status, nao_conforme_fmt)

        print(f"\n✅ Análise detalhada concluída! Relatório salvo em:\n   {output_path}")
    except PermissionError:
        print(f"\n❌ Erro: Não foi possível salvar o arquivo '{output_path}'.")
        print("⚠️  O arquivo pode estar aberto no Excel ou em outro programa.")
        print("💡 Solução: Feche o arquivo Excel e execute o script novamente.")
        raise
    except Exception as e:
        print(f"\n❌ Erro ao salvar o relatório: {e}")
        raise

def executar():
    print("="*70)
    print("🏥 ANÁLISE DE EMPRÉSTIMOS ENTRE HOSPITAIS")
    print("="*70)
    
    try:
        df_saida, df_entrada = carregar_dados()
        if df_saida is None or df_entrada is None:
            return

        # Permite ajustar o limiar de similaridade
        limiar = 0.75  # 75% de similaridade mínima
        print(f"\n⚙️  Limiar de similaridade: {limiar:.0%}")
        
        df_resultado = analisar_itens(df_saida, df_entrada, limiar)
        gerar_estatisticas(df_resultado)
        salvar_relatorio(df_resultado)
        
        print("\n" + "="*70)
        print("✅ PROCESSO CONCLUÍDO COM SUCESSO!")
        print("="*70)
        
    except Exception as e:
        # Em caso de erro, salvar em arquivo de log
        log_path = os.path.join(os.path.expanduser("~"), "Desktop", "projeto análise de empréstimo", "erro_log.txt")
        with open(log_path, "w", encoding="utf-8") as f:
            import traceback
            f.write(f"Erro durante a execução: {e}\n\n")
            f.write(traceback.format_exc())
        print(f"\n❌ Erro durante execução. Log salvo em: {log_path}")
        raise

if __name__ == "__main__":
    executar()