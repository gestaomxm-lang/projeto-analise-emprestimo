import imaplib
import email
import os
from datetime import datetime
from email.utils import parsedate_to_datetime
from email.header import decode_header

# CONFIGURAÇÕES DE E-MAIL
EMAIL = 'gestao_mxm@grupohospitalcasa.com.br'
SENHA = 'rjjdyumxwkqdjdnp'
IMAP_SERVER = 'imap.gmail.com'
RECEBEDOR = 'pedro.gomes@hospitaldecancer.com.br'

PASTA_DESTINO = os.path.join(os.path.expanduser("~"), "Desktop", "projeto análise de empréstimo")

def baixar_anexos():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, SENHA)
    mail.select('inbox')

    # Busca todos os e-mails do remetente, sem filtro de data
    status, mensagens = mail.search(None, f'(FROM "{RECEBEDOR}")')

    if status != 'OK' or not mensagens[0]:
        print("Nenhum e-mail encontrado do remetente.")
        mail.logout()
        return

    # Pega apenas o mais recente
    ids = mensagens[0].split()
    ultimo_email_id = ids[-1]

    for num in [ultimo_email_id]:
        status, dados = mail.fetch(num, '(RFC822)')
        if status != 'OK':
            continue

        msg = email.message_from_bytes(dados[0][1])

        for parte in msg.walk():
            if parte.get_content_maintype() == 'multipart':
                continue
            if parte.get('Content-Disposition') is None:
                continue

            nome_arquivo = parte.get_filename()
            if nome_arquivo and nome_arquivo.endswith('.xlsx'):
                nome_arquivo_lower = nome_arquivo.lower()

                if 'emprestimo_concedido_' in nome_arquivo_lower:
                    caminho = os.path.join(PASTA_DESTINO, 'emprestimo_concedido.xlsx')
                elif 'emprestimo_recebido_' in nome_arquivo_lower:
                    caminho = os.path.join(PASTA_DESTINO, 'emprestimo_recebido.xlsx')
                else:
                    continue  # ignora outros arquivos

                with open(caminho, 'wb') as f:
                    f.write(parte.get_payload(decode=True))
                print(f"Anexo salvo: {caminho}")

    mail.logout()

if __name__ == "__main__":
    baixar_anexos()
