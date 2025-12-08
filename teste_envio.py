import requests
import os

# Configurações do DigiSac
DIGISAC_API_URL = "https://la3redecasa.digisac.co/api/v1"
TOKEN = "68491917eb4c7e5b67d998d9da1590a89dbe80d1"
NUMERO_DESTINATARIO = "5521994140008"  # Número do WhatsApp (incluindo código do país)
CAMINHO_ARQUIVO = r"C:\Users\MEUCOMPUTADOR\Desktop\projeto análise de empréstimo\analise_materiais_hospitalares.xlsx"


import mimetypes

def upload_arquivo():
    """Faz o upload do arquivo Excel para o DigiSac e retorna a URL do arquivo."""
    url_upload = f"{DIGISAC_API_URL}/files"
    headers = {"Authorization": f"Bearer {TOKEN}"}

    # Forçar a extensão e definir o mimetype corretamente
    mimetype = mimetypes.guess_type(CAMINHO_ARQUIVO)[0] or "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    with open(CAMINHO_ARQUIVO, "rb") as file:
        files = {
            "file": (
                os.path.basename(CAMINHO_ARQUIVO),  # Nome do arquivo
                file,  # Arquivo aberto
                mimetype  # Tipo do arquivo (Excel)
            )
        }
        response = requests.post(url_upload, headers=headers, files=files)

    if response.status_code == 200:
        return response.json().get("url")
    else:
        print("Erro ao fazer upload do arquivo:", response.text)
        return None





def enviar_arquivo_whatsapp(url_arquivo):
    """Envia a planilha via WhatsApp usando a API do DigiSac."""
    url_mensagem = f"{DIGISAC_API_URL}/messages"
    headers = {"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"}
    
    dados = {
        "number": NUMERO_DESTINATARIO,
        "type": "file",
        "content": {
            "url": url_arquivo,
            "filename": os.path.basename(CAMINHO_ARQUIVO),
            "caption": "Segue o relatório atualizado."
        }
    }
    
    response = requests.post(url_mensagem, json=dados, headers=headers)
    if response.status_code == 200:
        print("✅ Arquivo enviado com sucesso!")
    else:
        print("❌ Erro ao enviar arquivo:", response.text)


def executar_processo():
    """Executa o processo completo de upload e envio da planilha."""
    print("📤 Fazendo upload do arquivo...")
    url_arquivo = upload_arquivo()
    
    if url_arquivo:
        print("✅ Upload concluído! Enviando para WhatsApp...")
        enviar_arquivo_whatsapp(url_arquivo)
    else:
        print("❌ Processo encerrado devido a erro no upload.")


if __name__ == "__main__":
    executar_processo()
