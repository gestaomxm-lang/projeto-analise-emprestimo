import pandas as pd

# Verifica se 5085232 existe no doc_index
df_entrada = pd.read_excel('emprestimo_recebido.xlsx')

# Extrai números dos documentos
import re
def extrair_numeros(documento):
    match = re.search(r'\d+', str(documento))
    return match.group() if match else ""

df_entrada['doc_num'] = df_entrada['documento'].apply(extrair_numeros)

# Cria índice
doc_index = {}
for idx, row in df_entrada.iterrows():
    doc = row['doc_num']
    if doc and doc != '':
        if doc not in doc_index:
            doc_index[doc] = []
        doc_index[doc].append(idx)

print(f"Total de documentos únicos no índice: {len(doc_index)}")
print(f"\n5085232 está no doc_index? {'5085232' in doc_index}")

if '5085232' in doc_index:
    print(f"Número de itens: {len(doc_index['5085232'])}")
    print("PROBLEMA: O documento existe no índice, então o flag não será ativado!")
else:
    print("✅ Documento não está no índice, flag deveria ser ativado")

# Mostra alguns documentos próximos
docs_numericos = sorted([int(d) for d in doc_index.keys() if d.isdigit()])
print(f"\nDocumentos próximos a 5085232:")
for d in docs_numericos:
    if 5085000 <= d <= 5086000:
        print(f"  {d}: {len(doc_index[str(d)])} itens")
