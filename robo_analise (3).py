import pandas as pd
import os
import psycopg2

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
    
    df_saida['data'] = pd.to_datetime(df_saida['data']).dt.strftime("%d/%m/%Y")
    df_entrada['data'] = pd.to_datetime(df_entrada['data']).dt.strftime("%d/%m/%Y")
    
    return df_saida, df_entrada

def analisar_conformidade(df_saida, df_entrada):
    """Executa a análise garantindo múltiplas entradas para um mesmo documento e identificando duplicidades."""
    analise_detalhada = []
    df_entrada['documento'] = df_entrada['documento'].astype(str).str.strip()
    df_saida['documento'] = df_saida['documento'].astype(str).str.strip()
    
    duplicatas = df_entrada['documento'].str.extract(r'(\d+)')[0].value_counts()
    documentos_duplicados = duplicatas[duplicatas > 10].index.tolist()
    
    for _, saida in df_saida.iterrows():
        doc_saida = saida['documento']
        unidade_origem = saida['unidade_origem']
        unidade_destino = saida['unidade_destino']
        data_saida = saida['data']
        valor_saida = saida['valor_total']
        especie_saida = saida['especie']
        
        entrada_match = df_entrada[df_entrada['documento'].str.contains(str(doc_saida), na=False, case=False)]
        
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
                
                if abs(diferenca_valor) > 1.00:
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

def salvar_relatorio(df, caminho):
    """Salva os relatórios com formatação aprimorada."""
    with pd.ExcelWriter(caminho, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Relatório', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Relatório']
        
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
    
    df_analise_materiais = analisar_conformidade(
        df_saida[df_saida['especie'] == "MATERIAIS HOSPITALARES"],
        df_entrada[df_entrada['especie'] == "MATERIAIS HOSPITALARES"]
    )
    
    df_analise_medicamentos = analisar_conformidade(
        df_saida[df_saida['especie'] == "MEDICAMENTOS HOSPITALARES"],
        df_entrada[df_entrada['especie'] == "MEDICAMENTOS HOSPITALARES"]
    )
    
    df_final = pd.concat([df_analise_materiais, df_analise_medicamentos], ignore_index=True)
    output_final = os.path.join(projeto_path, "analise_emprestimos.xlsx")
    salvar_relatorio(df_final, output_final)
    
    print(f"Análise concluída! Relatório salvo em {output_final}")

if __name__ == "__main__":
    executar_analise()

def inserir_dados_postgresql(df, conn_params):
    """Insere os dados do dataframe na tabela analise_emprestimos no PostgreSQL"""
    conn = psycopg2.connect(**conn_params)
    cursor = conn.cursor()

    insert_query = """
    INSERT INTO analise_emprestimos (
        data, unidade_origem, unidade_destino, valor, documento_saida, 
        documento_entrada, especie, conformidade, status, divergencias
    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """

    for _, row in df.iterrows():
        cursor.execute(insert_query, tuple(row))

    conn.commit()
    cursor.close()
    conn.close()
    print("Dados inseridos com sucesso!")

conn_params = {
    "dbname": "database",
    "user": "rhc",
    "password": "24732070",
    "host": "192.168.6.15",
    "port": "5432"
}

df_resultado = pd.read_excel("analise_materiais_hospitalares.xlsx")
inserir_dados_postgresql(df_resultado, conn_params)