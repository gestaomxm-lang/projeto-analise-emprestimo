import pandas as pd
import sys
sys.path.insert(0, '.')
from app_test_2 import analisar_itens

# Carrega os arquivos
df_saida = pd.read_excel('emprestimo_concedido.xlsx')
df_entrada = pd.read_excel('emprestimo_recebido.xlsx')

# Roda a análise completa
print("=== Rodando análise completa ===")
df_resultado, stats = analisar_itens(df_saida, df_entrada, limiar_similaridade=65)

print("\n=== COLUNAS DO RESULTADO ===")
for i, col in enumerate(df_resultado.columns):
    print(f"{i}: {col}")

# Filtra resultado para METFORMINA
resultado_metf = df_resultado[
    df_resultado['Produto (Saída)'].str.contains('METFORMINA', case=False, na=False)
]

print(f"\n=== TODAS AS METFORMINAS NO RESULTADO ({len(resultado_metf)} registros) ===")
colunas_importantes = [
    'Produto (Saída)', 'Qtd Saída', 'Documento (Saída)',
    'Produto (Entrada)', 'Qtd Entrada', 'Documento (Entrada)',
    'Diferença Qtd', 'Conforme Qtd', 'Detalhes Match'
]
colunas_disponiveis = [col for col in colunas_importantes if col in df_resultado.columns]
print(resultado_metf[colunas_disponiveis].to_string(index=False))

# Salva resultado completo
df_resultado.to_csv('debug_metformina_resultado.csv', index=False, encoding='utf-8-sig')
print("\n✅ Resultado completo salvo em: debug_metformina_resultado.csv")
