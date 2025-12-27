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
            
        print(f"📧 Encontrados {len(email_ids)} emails. Procurando anexos no mais recente...")
        
        # Itera do mais recente para o mais antigo (reversed)
        for email_id in reversed(email_ids):
            status, msg_data = mail.fetch(email_id, "(RFC822)")
            
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    subject, encoding = decode_header(msg["Subject"])[0]
                    if isinstance(subject, bytes):
                        subject = subject.decode(encoding if encoding else "utf-8")
                    
                    print(f"📩 Verificando email: {subject}")
                    
                    # Verifica preliminar de anexos antes de limpar a pasta
                    temp_attachments = []
                    for part in msg.walk():
                        if part.get_content_maintype() == "multipart": continue
                        if part.get("Content-Disposition") is None: continue
                        filename = part.get_filename()
                        if filename and (filename.endswith(".xlsx") or filename.endswith(".xls")):
                            temp_attachments.append(part)
                            
                    if len(temp_attachments) >= 1: # Aceita se tiver pelo menos 1, mas idealmente 2
                        print(f"✅ Encontrados {len(temp_attachments)} anexos Excel neste e-mail. Baixando...")
                        
                        # Limpa pasta apenas agora que garantimos que vamos baixar novos
                        os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
                        for f in os.listdir(DOWNLOAD_FOLDER):
                            try:
                                os.remove(os.path.join(DOWNLOAD_FOLDER, f))
                            except: pass
                            
                        attachments_downloaded = 0
                        for part in temp_attachments:
                            filename = part.get_filename()
                            if decode_header(filename)[0][1]:
                                filename = decode_header(filename)[0][0].decode(decode_header(filename)[0][1])
                            
                            filepath = os.path.join(DOWNLOAD_FOLDER, filename)
                            with open(filepath, "wb") as f:
                                f.write(part.get_payload(decode=True))
                            print(f"⬇️ Baixado: {filename}")
                            attachments_downloaded += 1
                        
                        print(f"🎉 Sucesso! {attachments_downloaded} arquivos atualizados.")
                        return True
                    else:
                        print("⚠️ Este email não contém arquivos Excel válidos. Tentando o próximo mais recente...")
        
        print("❌ Nenhum email com anexos Excel válidos encontrado nos últimos 5 dias.")
        return False

    except Exception as e:
        print(f"Erro durante processamento: {e}")
        return False
    finally:
        mail.close()
        mail.logout()

if __name__ == "__main__":
    download_daily_attachments()
