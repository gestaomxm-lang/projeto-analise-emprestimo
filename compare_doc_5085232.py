import pandas as pd

df_saida = pd.read_excel('emprestimo_concedido.xlsx')
df_entrada = pd.read_excel('emprestimo_recebido.xlsx')

print("=== DOCUMENTO 5085232 NA SAÍDA ===")
saida_5085232 = df_saida[df_saida['documento'] == 5085232]
print(f"Total: {len(saida_5085232)} itens\n")
print("Produtos:")
for idx, row in saida_5085232.iterrows():
    print(f"  - {row['ds_produto']}")

print("\n=== DOCUMENTO 5085232 NA ENTRADA ===")
entrada_5085232 = df_entrada[df_entrada['documento'] == 5085232]
print(f"Total: {len(entrada_5085232)} itens\n")
print("Produtos:")
for idx, row in entrada_5085232.iterrows():
    print(f"  - {row['ds_produto']}")

# Compara METFORMINA
metf_saida = saida_5085232[saida_5085232['ds_produto'].str.contains('METFORMINA', case=False, na=False)]
metf_entrada = entrada_5085232[entrada_5085232['ds_produto'].str.contains('METFORMINA', case=False, na=False)]

print(f"\n=== METFORMINA ===")
print(f"Na saída: {len(metf_saida)} item(s)")
if not metf_saida.empty:
    print(f"  Produto: {metf_saida.iloc[0]['ds_produto']}")
    print(f"  Quantidade: {metf_saida.iloc[0]['qt_entrada']}")

print(f"\nNa entrada: {len(metf_entrada)} item(s)")
if not metf_entrada.empty:
    print(f"  Produto: {metf_entrada.iloc[0]['ds_produto']}")
    print(f"  Quantidade: {metf_entrada.iloc[0]['qt_entrada']}")
else:
    print("  ❌ METFORMINA NÃO EXISTE NA ENTRADA!")
