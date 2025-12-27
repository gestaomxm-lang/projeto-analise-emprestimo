import smtplib
import imaplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os
import datetime
import json

# Configurações do "Banco de Dados" via Email
DB_SUBJECT_TAG = "[SYSTEM_DB_USERS_BACKUP]"
LOCAL_DB_FILE = "users.json"

# Credenciais (Pega as mesmas do download_gmail.py)
EMAIL_USER = "doc.analise.robo@gmail.com" # Hardcoded conforme padrão do projeto existente
# Na prática deveria vir de ENV, mas manteremos padrão para funcionar direto
EMAIL_PASS = os.getenv("GMAIL_APP_PASSWORD") 
# Se não estiver no ENV, tentar hardcoded (USER NÃO FORNECEU A SENHA AQUI, VAMOS TENTAR LER DO DOWNLOAD_GMAIL SE PRECISAR)
# Assumindo que o ambiente já tem ou o usuário vai rodar onde tem.

def get_credentials():
    # Tenta pegar variáveis de ambiente primeiro
    user = os.getenv("GMAIL_USER", "doc.analise.robo@gmail.com")
    password = os.getenv("GMAIL_APP_PASSWORD")
    return user, password

def sync_up():
    """Envia o arquivo users.json local para o email (Backup/Save Cloud)."""
    if not os.path.exists(LOCAL_DB_FILE):
        print("⚠️ Arquivo local não existe para upload.")
        return False

    user, password = get_credentials()
    if not user or not password:
        print("❌ Credenciais de e-mail não encontradas.")
        return False

    try:
        msg = MIMEMultipart()
        msg['From'] = user
        msg['To'] = user
        msg['Subject'] = f"{DB_SUBJECT_TAG} {datetime.datetime.now().isoformat()}"
        
        body = "Backup automático do banco de usuários."
        msg.attach(MIMEText(body, 'plain'))
        
        with open(LOCAL_DB_FILE, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {LOCAL_DB_FILE}")
        msg.attach(part)
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(user, password)
        text = msg.as_string()
        server.sendmail(user, user, text)
        server.quit()
        print("✅ [Cloud Sync] Banco de usuários enviado para nuvem (Email).")
        return True
    except Exception as e:
        print(f"❌ [Cloud Sync] Erro no upload: {e}")
        return False

def sync_down():
    """Baixa a versão mais recente do users.json do email (Load Cloud)."""
    user, password = get_credentials()
    if not user or not password:
        print("❌ Credenciais de e-mail não encontradas.")
        return False
        
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(user, password)
        mail.select("inbox")
        
        # Busca emails enviados por mim com o assunto específico
        status, messages = mail.search(None, f'(FROM "{user}" SUBJECT "{DB_SUBJECT_TAG}")')
        
        if status != "OK" or not messages[0]:
            print("⚠️ [Cloud Sync] Nenhum backup encontrado na nuvem.")
            return False
            
        # Pega o último ID (mais recente)
        latest_email_id = messages[0].split()[-1]
        
        status, data = mail.fetch(latest_email_id, "(RFC822)")
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)
        
        found_attachment = False
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
                
            filename = part.get_filename()
            if filename == LOCAL_DB_FILE:
                filepath = os.path.join(os.getcwd(), LOCAL_DB_FILE)
                with open(filepath, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                found_attachment = True
                print("✅ [Cloud Sync] Banco de usuários baixado da nuvem.")
                break
        
        mail.close()
        mail.logout()
        return found_attachment
        
    except Exception as e:
        print(f"❌ [Cloud Sync] Erro no download: {e}")
        return False

if __name__ == "__main__":
    # Teste rápido se rodado direto
    print("Testando Sync...")
    # sync_up()
