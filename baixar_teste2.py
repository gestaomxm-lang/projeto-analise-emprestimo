import imaplib
import email
import os
from email.header import decode_header

# CONFIGURAÇÕES DE CONTA
EMAIL = "gestao_mxm@grupohospitalcasa.com.br"
SENHA = "rjjdyumxwkqdjdnp"  # senha de app gerada no Gmail
IMAP_SERVER = "imap.gmail.com"
RECEBEDOR = "pedro.gomes@hospitaldecancer.com.br"

# LOCAL DE SALVAMENTO
PASTA_DESTINO = os.path.join(os.path.expanduser("~"), "Desktop", "projeto análise de empréstimo")

def baixar_anexos():
    """Baixa anexos .xlsx do último e-mail do remetente e salva como saida.xlsx e entrada.xlsx"""
    if not os.path.exists(PASTA_DESTINO):
        os.makedirs(PASTA_DESTINO)

    print("📬 Conectando ao servidor IMAP...")
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, SENHA)
    mail.select("inbox")

    # Buscar e-mails do remetente especificado
    status, mensagens = mail.search(None, f'FROM {RECEBEDOR}')
    if status != "OK" or not mensagens[0]:
        print("⚠️ Nenhum e-mail encontrado desse remetente.")
        mail.logout()
        return

    ids = mensagens[0].split()
    ultimo_email_id = ids[-1]

    print(f"📥 E-mail mais recente encontrado (ID {ultimo_email_id.decode()}), iniciando download...")

    status, dados = mail.fetch(ultimo_email_id, "(RFC822)")
    if status != "OK":
        print("❌ Falha ao buscar o e-mail.")
        mail.logout()
        return

    msg = email.message_from_bytes(dados[0][1])

    anexos_salvos = 0
    for parte in msg.walk():
        if parte.get_content_maintype() == "multipart":
            continue
        if parte.get("Content-Disposition") is None:
            continue

        nome_arquivo = parte.get_filename()
        if nome_arquivo:
            nome_arquivo, _ = decode_header(nome_arquivo)[0]
            if isinstance(nome_arquivo, bytes):
                nome_arquivo = nome_arquivo.decode(errors="ignore")

            nome_arquivo_lower = nome_arquivo.lower()

            # Identifica tipo do arquivo e define o nome final
            if "emprestimo_concedido" in nome_arquivo_lower:
                caminho = os.path.join(PASTA_DESTINO, "emprestimo_concedido.xlsx")
                tipo = "saída"
            elif "emprestimo_recebido" in nome_arquivo_lower:
                caminho = os.path.join(PASTA_DESTINO, "emprestimo_recebido.xlsx")
                tipo = "entrada"
            else:
                continue  # ignora outros arquivos

            with open(caminho, "wb") as f:
                f.write(parte.get_payload(decode=True))

            anexos_salvos += 1
            print(f"✅ Arquivo de {tipo} salvo: {caminho}")

    if anexos_salvos == 0:
        print("⚠️ Nenhum anexo .xlsx encontrado neste e-mail.")
    else:
        print(f"🎯 {anexos_salvos} arquivo(s) salvo(s) com sucesso em: {PASTA_DESTINO}")

    mail.logout()

if __name__ == "__main__":
    baixar_anexos()
