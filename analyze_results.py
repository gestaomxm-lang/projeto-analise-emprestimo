import pandas as pd

df = pd.read_csv('teste_correcao_resultado.csv')

print("=== DOCUMENTO 5085232 ===")
doc_5085232 = df[df['Documento'] == 5085232]
print(f"Total de itens: {len(doc_5085232)}")

# Encontra METFORMINA
metf = doc_5085232[doc_5085232['Produto (Saída)'].str.contains('METFORMINA', case=False, na=False)]

if not metf.empty:
    print(f"\n=== METFORMINA ===")
    print(f"Status: {metf['Status'].values[0]}")
    print(f"Tipo de Divergência: {metf['Tipo de Divergência'].values[0]}")
    
    if 'Não Conforme' in str(metf['Status'].values[0]):
        print("\n✅ CORREÇÃO BEM-SUCEDIDA!")
        print("METFORMINA do documento 5085232 agora está corretamente marcada como 'Não Encontrado'")
    else:
        print("\n❌ PROBLEMA PERSISTE")
        print(f"Produto Entrada: {metf['Produto (Entrada)'].values[0]}")
else:
    print("METFORMINA não encontrada no documento 5085232")

# Mostra todos os itens do documento 5085232
print(f"\n=== TODOS OS ITENS DO DOCUMENTO 5085232 ===")
nao_conforme = doc_5085232[doc_5085232['Status'].str.contains('Não Conforme', na=False)]
print(f"Total de 'Não Conforme': {len(nao_conforme)} de {len(doc_5085232)}")

if len(nao_conforme) == 21:
    print("✅ PERFEITO! Todos os 21 itens estão marcados como 'Não Conforme'")
elif len(nao_conforme) > 0:
    print(f"⚠️ Apenas {len(nao_conforme)} itens marcados como 'Não Conforme'")
    print("\nItens NÃO marcados como 'Não Conforme':")
    conformes = doc_5085232[~doc_5085232['Status'].str.contains('Não Conforme', na=False)]
    for idx, row in conformes.iterrows():
        print(f"  - {row['Produto (Saída)']} | Status: {row['Status']}")
else:
    print("❌ Nenhum item marcado como 'Não Conforme'")
