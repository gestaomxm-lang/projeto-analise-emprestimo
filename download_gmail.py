import imaplib
import email
from email.header import decode_header
import os
from datetime import datetime
import datetime as dt
import re

# Configurações - Substitua pelos seus dados ou use variáveis de ambiente
# Para gerar sennha de app: https://myaccount.google.com/apppasswords
# Configurações - Substitua pelos seus dados ou use variáveis de ambiente
# Para gerar sennha de app: https://myaccount.google.com/apppasswords
EMAIL_USER = os.environ.get("GMAIL_USER", "gestao_mxm@grupohospitalcasa.com.br")
# Remove espaços da senha caso venha copiada com eles
EMAIL_PASS = os.environ.get("GMAIL_APP_PASSWORD", "eprk lgzt jvqy mqht").replace(" ", "")
SEARCH_SENDER = os.environ.get("GMAIL_SENDER", "pedro.gomes@hospitaldecancer.com.br") 
SEARCH_SUBJECT = os.environ.get("GMAIL_SUBJECT", "") # Deixe vazio para ignorar
DOWNLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dados", "input")

def connect_gmail():
    """Conecta ao Gmail usando IMAP."""
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(EMAIL_USER, EMAIL_PASS)
        return mail
    except Exception as e:
        print(f"Erro ao conectar no Gmail: {e}")
        return None

def download_daily_attachments():
    """Busca e baixa os anexos do dia."""
    if EMAIL_USER == "seu_email@gmail.com":
        print("⚠️ Configure as credenciais no arquivo ou variáveis de ambiente!")
        return False

    mail = connect_gmail()
    if not mail:
        return False

    try:
        mail.select("inbox")

        # Data de hoje para filtro (IMAP exige formato específico: 26-Dec-2025)
        # MODIFICADO: Busca últimos 5 dias para garantir que encontre algo para teste
        date_str = (datetime.now() - dt.timedelta(days=5)).strftime("%d-%b-%Y")
        
        # Constrói a query de busca
        query = f'(SINCE "{date_str}")'
        if SEARCH_SENDER:
            query += f' FROM "{SEARCH_SENDER}"'
        if SEARCH_SUBJECT:
            query += f' SUBJECT "{SEARCH_SUBJECT}"'
            
        print(f"🔍 Buscando emails com query: {query}")
        status, messages = mail.search(None, query)
        
        if status != "OK":
            print("Erro na busca.")
            return False
            
        email_ids = messages[0].split()
        
        if not email_ids:
            print("❌ Nenhum email encontrado hoje com os critérios definidos.")
            return False
            
        print(f"📧 Encontrados {len(email_ids)} emails. Processando o mais recente...")
        
        # Pega o mais recente
        latest_email_id = email_ids[-1]
        status, msg_data = mail.fetch(latest_email_id, "(RFC822)")
        
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")
                
                print(f"📩 Processando email: {subject}")
                
                attachments_found = 0
                
                # Garante diretório
                os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
                
                # Limpa diretório antes de baixar novos (opcional, para garantir frescor)
                for f in os.listdir(DOWNLOAD_FOLDER):
                    os.remove(os.path.join(DOWNLOAD_FOLDER, f))
                
                for part in msg.walk():
                    if part.get_content_maintype() == "multipart":
                        continue
                    if part.get("Content-Disposition") is None:
                        continue
                        
                    filename = part.get_filename()
                    if filename:
                        # Decodifica nome se necessário
                        if decode_header(filename)[0][1]:
                            filename = decode_header(filename)[0][0].decode(decode_header(filename)[0][1])
                            
                        if filename.endswith(".xlsx") or filename.endswith(".xls"):
                            filepath = os.path.join(DOWNLOAD_FOLDER, filename)
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))
                            print(f"⬇️ Baixado: {filename}")
                            attachments_found += 1
                
                if attachments_found >= 2:
                    print(f"✅ Sucesso! {attachments_found} arquivos baixados em {DOWNLOAD_FOLDER}")
                    return True
                else:
                    print(f"⚠️ Atenção: Apenas {attachments_found} arquivos Excel encontrados (esperado 2).")
                    return attachments_found > 0

    except Exception as e:
        print(f"Erro durante processamento: {e}")
        return False
    finally:
        mail.close()
        mail.logout()

if __name__ == "__main__":
    download_daily_attachments()
