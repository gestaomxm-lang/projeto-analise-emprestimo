import pandas as pd

# Carrega os arquivos
df_saida = pd.read_excel('emprestimo_concedido.xlsx')
df_entrada = pd.read_excel('emprestimo_recebido.xlsx')

print("=== COLUNAS SAÍDA ===")
for i, col in enumerate(df_saida.columns):
    print(f"{i}: {col}")

print("\n=== COLUNAS ENTRADA ===")
for i, col in enumerate(df_entrada.columns):
    print(f"{i}: {col}")

# Procura por documento 5085232
doc_num = 5085232

# Tenta diferentes nomes de coluna
doc_col_saida = None
for col in df_saida.columns:
    if 'doc' in col.lower() or 'numero' in col.lower():
        doc_col_saida = col
        break

doc_col_entrada = None
for col in df_entrada.columns:
    if 'doc' in col.lower() or 'numero' in col.lower():
        doc_col_entrada = col
        break

prod_col_saida = None
for col in df_saida.columns:
    if 'prod' in col.lower() or 'desc' in col.lower():
        prod_col_saida = col
        break

prod_col_entrada = None
for col in df_entrada.columns:
    if 'prod' in col.lower() or 'desc' in col.lower():
        prod_col_entrada = col
        break

print(f"\n=== Colunas identificadas ===")
print(f"Doc Saída: {doc_col_saida}")
print(f"Doc Entrada: {doc_col_entrada}")
print(f"Produto Saída: {prod_col_saida}")
print(f"Produto Entrada: {prod_col_entrada}")

# Salva resultados em arquivo
with open('debug_doc_5085232.txt', 'w', encoding='utf-8') as f:
    f.write(f"=== SAÍDA - Documento {doc_num} ===\n")
    if doc_col_saida:
        saida_5085232 = df_saida[df_saida[doc_col_saida] == doc_num]
        f.write(f"Total de itens: {len(saida_5085232)}\n\n")
        if not saida_5085232.empty:
            f.write(saida_5085232.to_string())
        else:
            f.write("Nenhum registro encontrado\n")
    
    f.write(f"\n\n=== ENTRADA - Documento {doc_num} ===\n")
    if doc_col_entrada:
        entrada_5085232 = df_entrada[df_entrada[doc_col_entrada] == doc_num]
        f.write(f"Total de itens: {len(entrada_5085232)}\n\n")
        if not entrada_5085232.empty:
            f.write(entrada_5085232.to_string())
        else:
            f.write("Nenhum registro encontrado\n")
    
    # Procura METFORMINA
    f.write("\n\n=== METFORMINA na SAÍDA ===\n")
    if prod_col_saida:
        metf_saida = df_saida[df_saida[prod_col_saida].astype(str).str.contains('METFORMINA', case=False, na=False)]
        f.write(f"Total: {len(metf_saida)} registros\n\n")
        for idx, row in metf_saida.iterrows():
            f.write(f"Doc: {row[doc_col_saida]} | Produto: {row[prod_col_saida]}\n")
    
    f.write("\n\n=== METFORMINA na ENTRADA ===\n")
    if prod_col_entrada:
        metf_entrada = df_entrada[df_entrada[prod_col_entrada].astype(str).str.contains('METFORMINA', case=False, na=False)]
        f.write(f"Total: {len(metf_entrada)} registros\n\n")
        for idx, row in metf_entrada.iterrows():
            f.write(f"Doc: {row[doc_col_entrada]} | Produto: {row[prod_col_entrada]}\n")

print("\n✅ Resultados salvos em: debug_doc_5085232.txt")
