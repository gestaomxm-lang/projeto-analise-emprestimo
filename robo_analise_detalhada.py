import pandas as pd
import os
import re
import sys

def extrair_numeros(documento):
    """Extrai a primeira sequência numérica encontrada no documento."""
    match = re.search(r'\d+', str(documento))
    return match.group() if match else ""

def _parse_date_column(series):
    """Converte a coluna de datas para datetime."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_datetime(series, unit='d', origin='1899-12-30', errors='coerce')
    return pd.to_datetime(series, dayfirst=True, errors='coerce')

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

def analisar_itens(df_saida, df_entrada):
    """Compara item a item entre saída e entrada."""
    analise = []

    # Normaliza colunas e remove espaços extras
    for df in [df_saida, df_entrada]:
        df['documento'] = df['documento'].astype(str).str.strip()
        df['ds_produto'] = df['ds_produto'].astype(str).str.strip()
        df['unidade_origem'] = df['unidade_origem'].astype(str).str.strip()
        df['unidade_destino'] = df['unidade_destino'].astype(str).str.strip()

    # Loop por cada item de saída
    for _, row_s in df_saida.iterrows():
        doc_num = extrair_numeros(row_s['documento'])
        produto_s = row_s['ds_produto']
        valor_s = float(row_s['valor_total'])
        especie = row_s.get('especie', '')
        origem_s = row_s['unidade_origem']
        destino_s = row_s['unidade_destino']
        data_s = row_s['data']

        # Tenta localizar o mesmo documento e produto na entrada
        match = df_entrada[
            (df_entrada['ds_produto'].str.lower() == produto_s.lower()) &
            (df_entrada['documento'].apply(lambda x: extrair_numeros(x)) == doc_num)
        ]

        if not match.empty:
            for _, row_e in match.iterrows():
                valor_e = float(row_e['valor_total'])
                diferenca = round(valor_s - valor_e, 2)
                conforme = abs(diferenca) <= 10

                if conforme:
                    status = "Conforme"
                    tipo_div = "-"
                else:
                    status = "Não Conforme"
                    tipo_div = "Valor divergente"

                analise.append([
                    data_s,
                    origem_s,
                    destino_s,
                    doc_num,
                    produto_s,
                    especie,
                    valor_s,
                    valor_e,
                    diferenca,
                    status,
                    tipo_div
                ])
        else:
            analise.append([
                data_s,
                origem_s,
                destino_s,
                doc_num,
                produto_s,
                especie,
                valor_s,
                None,
                None,
                "Não Conforme",
                "Item não encontrado na entrada"
            ])

    # Agora procura itens recebidos sem saída correspondente
    for _, row_e in df_entrada.iterrows():
        doc_num_e = extrair_numeros(row_e['documento'])
        produto_e = row_e['ds_produto']

        match_saida = df_saida[
            (df_saida['ds_produto'].str.lower() == produto_e.lower()) &
            (df_saida['documento'].apply(lambda x: extrair_numeros(x)) == doc_num_e)
        ]

        if match_saida.empty:
            analise.append([
                row_e['data'],
                row_e['unidade_origem'],
                row_e['unidade_destino'],
                doc_num_e,
                produto_e,
                row_e.get('especie', ''),
                None,
                float(row_e['valor_total']),
                None,
                "Não Conforme",
                "Item não encontrado na saída"
            ])

    df_resultado = pd.DataFrame(analise, columns=[
        "Data", "Unidade Origem", "Unidade Destino", "Documento",
        "Produto", "Espécie", "Valor Saída (R$)", "Valor Entrada (R$)",
        "Diferença (R$)", "Status", "Tipo de Divergência"
    ])

    return df_resultado

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

            money_fmt = workbook.add_format({'num_format': 'R$ #,##0.00'})
            date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy'})

            ws.set_column('A:A', 12, date_fmt)
            ws.set_column('B:D', 25)
            ws.set_column('E:E', 60)
            ws.set_column('F:F', 25)
            ws.set_column('G:I', 18, money_fmt)
            ws.set_column('J:K', 30)

        print(f"✅ Análise detalhada concluída! Relatório salvo em:\n{output_path}")
    except PermissionError:
        print(f"❌ Erro: Não foi possível salvar o arquivo '{output_path}'.")
        print("⚠️  O arquivo pode estar aberto no Excel ou em outro programa.")
        print("💡 Solução: Feche o arquivo Excel e execute o script novamente.")
        raise
    except Exception as e:
        print(f"❌ Erro ao salvar o relatório: {e}")
        raise

def executar():
    try:
        df_saida, df_entrada = carregar_dados()
        if df_saida is None or df_entrada is None:
            return

        df_resultado = analisar_itens(df_saida, df_entrada)
        salvar_relatorio(df_resultado)
    except Exception as e:
        # Em caso de erro, salvar em arquivo de log
        log_path = os.path.join(os.path.expanduser("~"), "Desktop", "projeto análise de empréstimo", "erro_log.txt")
        with open(log_path, "w", encoding="utf-8") as f:
            import traceback
            f.write(f"Erro durante a execução: {e}\n\n")
            f.write(traceback.format_exc())

if __name__ == "__main__":
    executar()
    