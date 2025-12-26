import pandas as pd
import os
import glob
import pickle
import sys
from datetime import datetime
import analise_core
import download_gmail

# Diretórios
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "dados", "input")
DATA_DIR = os.path.join(BASE_DIR, "dados")
RESULT_FILE = os.path.join(DATA_DIR, "resultado_diario.pkl")
METADATA_FILE = os.path.join(DATA_DIR, "resultado_diario_metadata.json")

def pontuar_arquivo(df, nome_arquivo, termos_saida, termos_entrada):
    """Identifica se é arquivo de saída ou entrada baseado no nome."""
    nome_arquivo = nome_arquivo.lower()
    score_saida = sum(1 for t in termos_saida if t in nome_arquivo)
    score_entrada = sum(1 for t in termos_entrada if t in nome_arquivo)
    return score_saida, score_entrada

def executar_fluxo_diario(baixar_email=True):
    print(f"=== Iniciando Fluxo Diário: {datetime.now()} ===")
    
    # 1. Baixar Arquivos (Etapa 06:30)
    if baixar_email:
        print(">> Etapa 1: Baixando arquivos do Gmail...")
        sucesso_download = download_gmail.download_daily_attachments()
        if not sucesso_download:
            print("❌ Falha ao baixar arquivos ou nenhum arquivo encontrado. Abortando.")
            # Aqui entraria envio de alerta de erro
            return False
    else:
        print(">> Etapa 1: Pular download (usando arquivos locais)...")

    # 2. Processar Arquivos (Etapa 07:00)
    print(">> Etapa 2: Processando arquivos...")
    
    arquivos = glob.glob(os.path.join(INPUT_DIR, "*.xls*"))
    if len(arquivos) < 2:
        print(f"❌ Número insuficiente de arquivos em {INPUT_DIR}. Encontrados: {len(arquivos)}")
        return False
        
    arquivos_saida = []
    arquivos_entrada = []
    
    for arq_path in arquivos:
        try:
            nome_arquivo = os.path.basename(arq_path)
            print(f"   Lendo: {nome_arquivo}")
            df_temp = pd.read_excel(arq_path)
            
            s_saida, s_entrada = pontuar_arquivo(df_temp, nome_arquivo, ['saida', 'concedido', 'envio'], ['entrada', 'recebido'])
            
            if s_saida > s_entrada:
                arquivos_saida.append((nome_arquivo, df_temp))
            else:
                arquivos_entrada.append((nome_arquivo, df_temp))
        except Exception as e:
            print(f"   Erro ao ler {arq_path}: {e}")

    if not arquivos_saida or not arquivos_entrada:
        print("❌ Não foi possível identificar pares de Saída/Entrada.")
        return False

    # Consolidação
    print(f"   Identificados: {len(arquivos_saida)} Saída, {len(arquivos_entrada)} Entrada")
    
    df_saida = pd.concat([df for _, df in arquivos_saida], ignore_index=True)
    df_entrada = pd.concat([df for _, df in arquivos_entrada], ignore_index=True)
    
    nome_saida_consol = ", ".join([n for n, _ in arquivos_saida])
    nome_entrada_consol = ", ".join([n for n, _ in arquivos_entrada])
    
    # Preparação usando o Core
    print("   Preparando DataFrames...")
    df_saida = analise_core.preparar_dataframe(df_saida)
    df_entrada = analise_core.preparar_dataframe(df_entrada)
    
    # Execução da Análise
    print("   Executando algoritmo de análise...")
    
    def progress_wrapper(p, msg):
        print(f"   [{p*100:.0f}%] {msg}")
        
    df_resultado, stats = analise_core.analisar_itens(df_saida, df_entrada, progress_callback=progress_wrapper)
    
    # 3. Salvar Resultados para o Dashboard
    print(">> Etapa 3: Salvando resultados...")
    os.makedirs(DATA_DIR, exist_ok=True)
    
    # Salva pickle para carregamento rápido no Streamlit
    try:
        with open(RESULT_FILE, 'wb') as f:
            pickle.dump({
                'df': df_resultado,
                'metadata': {
                    'arquivo_saida': nome_saida_consol,
                    'arquivo_entrada': nome_entrada_consol,
                    'data_processamento': datetime.now(),
                    'stats': stats
                }
            }, f)
        print(f"✅ Resultado salvo em: {RESULT_FILE}")
    except Exception as e:
        print(f"❌ Erro ao salvar resultado: {e}")
        return False
        
    # (Opcional) Salvar histórico CSV também
    # analise_3.0.py tem função de salvar histórico, podemos replicar ou importar se quisermos persistência de longo prazo
    
    print("=== Processo Concluído com Sucesso ===")
    return True

if __name__ == "__main__":
    # Se passar argumento "download", baixa. Se passar "no-download", só processa.
    do_download = True
    if len(sys.argv) > 1 and sys.argv[1] == "no-download":
        do_download = False
        
    executar_fluxo_diario(baixar_email=do_download)
