import os
import yagmail

# Configurações
EMAIL = 'gestao_mxm@grupohospitalcasa.com.br'
SENHA = 'rjjdyumxwkqdjdnp'  # mesma senha de app
DESTINATARIOS = ['gerente.farmacia@hospitalcasa.com.br', "gestao_mxm@grupohospitalcasa.com.br"]
ARQUIVO = os.path.join(os.path.expanduser("~"), "Desktop", "projeto análise de empréstimo", "analise_emprestimos.xlsx")

def enviar_email():
    if not os.path.exists(ARQUIVO):
        print("Arquivo de análise não encontrado.")
        return

    yag = yagmail.SMTP(EMAIL, SENHA)
    assunto = "📊 Relatório Diário - Análise de Empréstimos"
    corpo = "Segue em anexo o relatório atualizado de movimentações de empréstimos."

    yag.send(
        to=DESTINATARIOS,
        subject=assunto,
        contents=corpo,
        attachments=ARQUIVO
    )
    print("E-mail enviado com sucesso!")

if __name__ == "__main__":
    enviar_email()
