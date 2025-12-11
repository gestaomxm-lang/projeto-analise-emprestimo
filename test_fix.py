import pandas as pd
import sys
sys.path.insert(0, '.')
from app_test_2 import analisar_itens

print("=== Testando correção do cross-document matching ===\n")

# Carrega os arquivos
df_saida = pd.read_excel('emprestimo_concedido.xlsx')
df_entrada = pd.read_excel('emprestimo_recebido.xlsx')

print(f"Total de itens na saída: {len(df_saida)}")
print(f"Total de itens na entrada: {len(df_entrada)}\n")

# Roda a análise completa
print("Rodando análise...")
df_resultado, stats = analisar_itens(df_saida, df_entrada, limiar_similaridade=65)

print(f"\n=== Estatísticas Gerais ===")
print(f"Conformes: {stats['conformes']}")
print(f"Não Conformes: {stats['nao_conformes']}")
print(f"Não Encontrados: {stats['nao_encontrados']}")

# Filtra resultado para documento 5085232
resultado_5085232 = df_resultado[df_resultado['Data'].apply(lambda x: True)]  # Pega todas as linhas primeiro
# Encontra a coluna correta de documento
doc_col = None
for col in df_resultado.columns:
    if 'documento' in col.lower() and 'sa' in col.lower():
        doc_col = col
        break

if doc_col:
    resultado_5085232 = df_resultado[df_resultado[doc_col] == 5085232]
    print(f"\n=== Documento 5085232 ({len(resultado_5085232)} itens) ===")
    
    # Conta quantos são "Não Encontrados"
    nao_encontrados = resultado_5085232[resultado_5085232['Status'].str.contains('Não Conforme', na=False)]
    print(f"Itens marcados como 'Não Conforme': {len(nao_encontrados)}")
    
    # Verifica se METFORMINA está marcada corretamente
    metformina = resultado_5085232[resultado_5085232['Produto (Saída)'].str.contains('METFORMINA', case=False, na=False)]
    if not metformina.empty:
        print(f"\n=== METFORMINA no documento 5085232 ===")
        print(f"Status: {metformina['Status'].values[0]}")
        print(f"Tipo de Divergência: {metformina['Tipo de Divergência'].values[0]}")
        
        if 'Não Conforme' in metformina['Status'].values[0]:
            print("✅ CORREÇÃO BEM-SUCEDIDA! METFORMINA agora está marcada como 'Não Encontrado'")
        else:
            print("❌ PROBLEMA: METFORMINA ainda está sendo matchada incorretamente")
            print(f"Produto Entrada: {metformina['Produto (Entrada)'].values[0]}")
            if doc_col:
                entrada_doc_col = doc_col.replace('Saída', 'Entrada')
                if entrada_doc_col in metformina.columns:
                    print(f"Documento Entrada: {metformina[entrada_doc_col].values[0]}")
else:
    print("Não foi possível encontrar a coluna de documento")

# Salva resultado
df_resultado.to_csv('teste_correcao_resultado.csv', index=False, encoding='utf-8-sig')
print("\n✅ Resultado completo salvo em: teste_correcao_resultado.csv")
