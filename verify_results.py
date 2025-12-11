import pandas as pd
import os
from datetime import datetime

# Verifica timestamp do arquivo
csv_path = 'teste_correcao_resultado.csv'
if os.path.exists(csv_path):
    mod_time = os.path.getmtime(csv_path)
    mod_datetime = datetime.fromtimestamp(mod_time)
    print(f"Arquivo CSV modificado em: {mod_datetime}")
    print(f"Hora atual: {datetime.now()}\n")

df = pd.read_csv(csv_path)

# Filtra documento 5085232
doc_5085232 = df[df['Documento'] == 5085232]
print(f"=== DOCUMENTO 5085232 ({len(doc_5085232)} itens) ===\n")

# Verifica status geral
status_counts = doc_5085232['Status'].value_counts()
print("Status:")
for status, count in status_counts.items():
    print(f"  {status}: {count}")

# Procura METFORMINA
metf = doc_5085232[doc_5085232['Produto (Saída)'].str.contains('METFORMINA', case=False, na=False)]

if not metf.empty:
    print(f"\n=== METFORMINA ===")
    print(f"Status: {metf['Status'].values[0]}")
    print(f"Tipo Div: {metf['Tipo de Divergência'].values[0]}")
    print(f"Produto Entrada: {metf['Produto (Entrada)'].values[0]}")
    
    if 'Não Conforme' in str(metf['Status'].values[0]) and str(metf['Produto (Entrada)'].values[0]) == '-':
        print("\n✅ CORREÇÃO BEM-SUCEDIDA!")
        print("METFORMINA está marcada como 'Não Encontrado' (Produto Entrada = '-')")
    elif 'Não Conforme' in str(metf['Status'].values[0]):
        print("\n✅ PARCIALMENTE CORRETO")
        print("Status é 'Não Conforme', mas ainda mostra um produto de entrada")
    else:
        print("\n❌ PROBLEMA PERSISTE")
        print("METFORMINA ainda está sendo matchada incorretamente")
else:
    print("\nMETFORMINA não encontrada no resultado")
