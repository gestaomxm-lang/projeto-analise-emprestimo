import pandas as pd
import os
import re

def extrair_numeros(documento):
    """Extrai a primeira sequência numérica encontrada no documento."""
    match = re.search(r'\d+', str(documento))
    return match.group() if match else ""

def _parse_date_column(series):
    """Converte a coluna de datas para datetime, tratando strings (dd/mm/yyyy)
    e valores numéricos vindos como serial do Excel.
    """
    # Se vier como numérico (serial do Excel)
    if pd.api.types.is_numeric_dtype(series):
        dt = pd.to_datetime(series, unit='d', origin='1899-12-30', errors='coerce')
    else:
        # Trata strings variadas, priorizando dia/mês/ano
        dt = pd.to_datetime(series, dayfirst=True, errors='coerce', infer_datetime_format=True)
    return dt

def carregar_dados():
    """Busca automaticamente os arquivos de saída e entrada na pasta 'projeto análise de empréstimo' no Desktop."""
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    projeto_path = os.path.join(desktop_path, "projeto análise de empréstimo")
    
    saida_path = os.path.join(projeto_path, "emprestimo_concedido.xlsx")
    entrada_path = os.path.join(projeto_path, "emprestimo_recebido.xlsx")
    
    if not os.path.exists(saida_path) or not os.path.exists(entrada_path):
        print("Erro: Arquivos emprestimo_concedido.xlsx e emprestimo_recebido.xlsx não encontrados na pasta 'projeto análise de empréstimo' no Desktop.")
        return None, None
    
    df_saida = pd.read_excel(saida_path)
    df_entrada = pd.read_excel(entrada_path)
    
    df_saida.rename(columns={'Unidade Destino': 'unidade_destino', 'Unidade Origem': 'unidade_origem'}, inplace=True)
    df_entrada.rename(columns={'Unidade Destino': 'unidade_destino', 'Unidade Origem': 'unidade_origem'}, inplace=True)
    
    # Converte as colunas de data para datetime sem formatar para string,
    # deixando a formatação para o ExcelWriter.
    df_saida['data'] = _parse_date_column(df_saida['data'])
    df_entrada['data'] = _parse_date_column(df_entrada['data'])
    
    return df_saida, df_entrada

def analisar_conformidade(df_saida, df_entrada):
    """Executa a análise garantindo múltiplas entradas para um mesmo documento e identificando duplicidades a partir da saída."""
    analise_detalhada = []
    df_entrada['documento'] = df_entrada['documento'].astype(str).str.strip()
    df_saida['documento'] = df_saida['documento'].astype(str).str.strip()
    
    duplicatas = df_entrada['documento'].str.extract(r'(\d+)')[0].value_counts()
    documentos_duplicados = duplicatas[duplicatas > 1].index.tolist()
    
    for _, saida in df_saida.iterrows():
        doc_saida = saida['documento']
        unidade_origem = saida['unidade_origem']
        unidade_destino = saida['unidade_destino']
        data_saida = saida['data']
        valor_saida = saida['valor_total']
        especie_saida = saida['especie']
        
        num_doc_saida = extrair_numeros(doc_saida)
        entrada_match = df_entrada[df_entrada['documento'].apply(lambda x: extrair_numeros(x) == num_doc_saida)]
        
        if not entrada_match.empty:
            for _, entrada in entrada_match.iterrows():
                conformidade = "Conforme"
                status = "Entrada confirmada"
                divergencias = []
                
                doc_base = entrada['documento'].split('-')[0]
                if doc_base in documentos_duplicados:
                    conformidade = "Não Conforme"
                    status = "Entrada duplicada"
                    divergencias.append("Documento de entrada duplicado")

                if unidade_origem != entrada['unidade_origem']:
                    conformidade = "Não Conforme"
                    divergencias.append("Unidade de Origem divergente")

                if unidade_destino != entrada['unidade_destino']:
                    conformidade = "Não Conforme"
                    divergencias.append("Unidade de Destino divergente")

                valor_entrada = entrada['valor_total']
                diferenca_valor = round(valor_saida - valor_entrada, 2)
                
                if abs(diferenca_valor) > 10.00:
                    conformidade = "Não Conforme"
                    divergencias.append(f"Valor divergente (Saída: R${valor_saida:.2f} | Entrada: R${valor_entrada:.2f} | Diferença: R${diferenca_valor:.2f})")
                
                if conformidade == "Não Conforme":
                    status = "Divergências encontradas"

                analise_detalhada.append([
                    data_saida, unidade_origem, unidade_destino, valor_saida,
                    doc_saida, entrada['documento'], especie_saida, conformidade,
                    status, ", ".join(divergencias) if divergencias else "-"
                ])
        else:
            analise_detalhada.append([
                data_saida, unidade_origem, unidade_destino, valor_saida,
                doc_saida, "", especie_saida, "Não Conforme",
                "Saída sem entrada correspondente", "Entrada não encontrada"
            ])

    df_resultado = pd.DataFrame(analise_detalhada, columns=[
        'Data', 'Unidade Origem', 'Unidade Destino', 'Valor',
        'Documento Saída', 'Documento Entrada', 'Espécie', 'Conformidade', 'Status', 'Divergências'
    ])
    
    return df_resultado

def analisar_entradas_sem_saida(df_saida, df_entrada):
    """Realiza a análise de entradas sem saídas correspondentes.
       O relatório segue a mesma estrutura de colunas para possibilitar a consolidação."""
    analise_detalhada = []
    df_entrada['documento'] = df_entrada['documento'].astype(str).str.strip()
    df_saida['documento'] = df_saida['documento'].astype(str).str.strip()
    
    for _, entrada in df_entrada.iterrows():
        doc_entrada = entrada['documento']
        unidade_origem = entrada['unidade_origem']
        unidade_destino = entrada['unidade_destino']
        data_entrada = entrada['data']
        valor_entrada = entrada['valor_total']
        especie_entrada = entrada['especie']
        
        num_doc_entrada = extrair_numeros(doc_entrada)
        saida_match = df_saida[df_saida['documento'].apply(lambda x: extrair_numeros(x) == num_doc_entrada)]
        
        if saida_match.empty:
            # Para manter a consistência, 'Documento Saída' ficará vazio e 'Documento Entrada' conterá o valor
            analise_detalhada.append([
                data_entrada, unidade_origem, unidade_destino, valor_entrada,
                "", entrada['documento'], especie_entrada, "Não Conforme",
                "Entrada sem saída correspondente", "Saída não encontrada"
            ])
    
    df_resultado = pd.DataFrame(analise_detalhada, columns=[
        'Data', 'Unidade Origem', 'Unidade Destino', 'Valor',
        'Documento Saída', 'Documento Entrada', 'Espécie', 'Conformidade', 'Status', 'Divergências'
    ])
    
    return df_resultado

def salvar_relatorio(df, caminho):
    """Salva o relatório com formatação aprimorada em um único arquivo Excel."""
    with pd.ExcelWriter(caminho, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Relatório Completo', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Relatório Completo']
        
        format1 = workbook.add_format({'num_format': 'dd/mm/yyyy'})
        format2 = workbook.add_format({'num_format': 'R$#,##0.00'})
        
        worksheet.set_column('A:A', 12, format1)
        worksheet.set_column('B:D', 20)
        worksheet.set_column('E:F', 25)
        worksheet.set_column('G:G', 20)
        worksheet.set_column('H:H', 15)
        worksheet.set_column('I:I', 25)
        worksheet.set_column('J:J', 50)
        worksheet.set_column('D:D', 15, format2)

def executar_analise():
    df_saida, df_entrada = carregar_dados()
    if df_saida is None or df_entrada is None:
        return
    
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    projeto_path = os.path.join(desktop_path, "projeto análise de empréstimo")
    
    # Análise de conformidade (saídas com entrada correspondente) para ambas as espécies
    df_analise_materiais = analisar_conformidade(
        df_saida[df_saida['especie'] == "MATERIAIS HOSPITALARES"],
        df_entrada[df_entrada['especie'] == "MATERIAIS HOSPITALARES"]
    )
    df_analise_medicamentos = analisar_conformidade(
        df_saida[df_saida['especie'] == "MEDICAMENTOS HOSPITALARES"],
        df_entrada[df_entrada['especie'] == "MEDICAMENTOS HOSPITALARES"]
    )
    
    # Análise de entradas sem saídas correspondentes para ambas as espécies
    df_analise_entradas_materiais = analisar_entradas_sem_saida(
        df_saida[df_saida['especie'] == "MATERIAIS HOSPITALARES"],
        df_entrada[df_entrada['especie'] == "MATERIAIS HOSPITALARES"]
    )
    df_analise_entradas_medicamentos = analisar_entradas_sem_saida(
        df_saida[df_saida['especie'] == "MEDICAMENTOS HOSPITALARES"],
        df_entrada[df_entrada['especie'] == "MEDICAMENTOS HOSPITALARES"]
    )
    
    # Consolida todas as análises em um único DataFrame
    df_final = pd.concat([
        df_analise_materiais, 
        df_analise_medicamentos, 
        df_analise_entradas_materiais, 
        df_analise_entradas_medicamentos
    ], ignore_index=True)
    
    output_final = os.path.join(projeto_path, "analise_emprestimos.xlsx")
    salvar_relatorio(df_final, output_final)
    
    print(f"Análise concluída! Relatório salvo em {output_final}")

if __name__ == "__main__":
    executar_analise()
