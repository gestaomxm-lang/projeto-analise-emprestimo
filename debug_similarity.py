import pandas as pd
import sys
sys.path.insert(0, '.')
from app_test_2 import extrair_componentes_produto, calcular_similaridade_precalc

# Carrega dados
df_saida = pd.read_excel('emprestimo_concedido.xlsx')
df_entrada = pd.read_excel('emprestimo_recebido.xlsx')

# Pega METFORMINA da saída
metf_saida = df_saida[(df_saida['documento'] == 5085232) & 
                       (df_saida['ds_produto'].str.contains('METFORMINA', case=False, na=False))].iloc[0]

# Pega todos os itens do doc 5085232 na entrada
entrada_5085232 = df_entrada[df_entrada['documento'] == 5085232]

print("=== METFORMINA (SAÍDA) ===")
print(f"Produto: {metf_saida['ds_produto']}")
print(f"Quantidade: {metf_saida['qt_entrada']}")

comp_metf = extrair_componentes_produto(metf_saida['ds_produto'])
print(f"\nComponentes extraídos:")
print(f"  Normalizado: {comp_metf['normalizado']}")
print(f"  Princípio Ativo: {comp_metf['principio_ativo']}")
print(f"  Concentração: {comp_metf['concentracao']}")

print(f"\n=== CANDIDATOS NO DOC 5085232 (ENTRADA) ===")
print(f"Total de itens: {len(entrada_5085232)}\n")

# Testa similaridade com cada item
scores = []
for idx, row_e in entrada_5085232.iterrows():
    comp_e = extrair_componentes_produto(row_e['ds_produto'])
    score, detalhes = calcular_similaridade_precalc(comp_metf, comp_e, ignore_penalties=True)
    qtd_e = float(row_e['qt_entrada'])
    qtd_match = abs(qtd_e - metf_saida['qt_entrada']) < 0.01
    limiar = 70 if qtd_match else 85
    
    scores.append({
        'produto': row_e['ds_produto'],
        'qtd': qtd_e,
        'qtd_match': qtd_match,
        'limiar': limiar,
        'score': score,
        'aceito': score >= limiar,
        'detalhes': detalhes
    })

# Ordena por score
scores.sort(key=lambda x: x['score'], reverse=True)

print("Top 5 candidatos por similaridade:")
for i, s in enumerate(scores[:5], 1):
    status = "✅ ACEITO" if s['aceito'] else "❌ REJEITADO"
    print(f"\n{i}. {status} (Score: {s['score']:.0f}% >= {s['limiar']}%)")
    print(f"   Produto: {s['produto'][:60]}")
    print(f"   Qtd: {s['qtd']} (Match exato: {s['qtd_match']})")
    print(f"   Detalhes: {s['detalhes']}")

# Verifica se algum foi aceito
aceitos = [s for s in scores if s['aceito']]
if aceitos:
    print(f"\n❌ PROBLEMA: {len(aceitos)} item(ns) foi(ram) aceito(s)!")
    print("METFORMINA será matchada incorretamente.")
else:
    print(f"\n✅ SUCESSO: Nenhum item foi aceito!")
    print("METFORMINA será corretamente marcada como 'Não Encontrado'.")
