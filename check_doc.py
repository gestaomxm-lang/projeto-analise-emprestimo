import pandas as pd

df = pd.read_csv('teste_correcao_resultado.csv')

print("=== DOCUMENTO 5085232 ===")
doc_5085232 = df[df['Documento'] == 5085232]
print(f"Total de itens: {len(doc_5085232)}\n")

# Agrupa por status
status_counts = doc_5085232['Status'].value_counts()
print("Distribuição de Status:")
for status, count in status_counts.items():
    print(f"  {status}: {count}")

# Mostra os primeiros 5 itens
print("\n=== PRIMEIROS 5 ITENS ===")
for idx, row in doc_5085232.head(5).iterrows():
    print(f"{row['Produto (Saída)'][:50]:<50} | {row['Status']}")

# Verifica METFORMINA especificamente
metf = doc_5085232[doc_5085232['Produto (Saída)'].str.contains('METFORMINA', case=False, na=False)]
if not metf.empty:
    print(f"\n=== METFORMINA ===")
    print(f"Produto: {metf['Produto (Saída)'].values[0]}")
    print(f"Status: {metf['Status'].values[0]}")
    print(f"Tipo Div: {metf['Tipo de Divergência'].values[0]}")
    print(f"Produto Entrada: {metf['Produto (Entrada)'].values[0]}")
    
    # Verifica se tem documento de entrada
    if 'Documento (Entrada)' in metf.columns:
        print(f"Doc Entrada: {metf['Documento (Entrada)'].values[0]}")
