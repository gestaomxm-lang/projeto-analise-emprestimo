import subprocess
import time
import sys
import logging
import os

# Configurações de log
logging.basicConfig(
    filename=os.path.join(os.path.dirname(__file__), 'wrapper.log'),
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

def run_script(path):
    """Executa um executável e registra saída/erros."""
    try:
        logging.info(f"Iniciando: {path}")
        result = subprocess.run([path], check=True, capture_output=True, text=True)
        logging.info(f"Saída de {os.path.basename(path)}:\n{result.stdout}")
    except subprocess.CalledProcessError as e:
        logging.error(f"Erro em {os.path.basename(path)}: {e}\nSTDERR:\n{e.stderr}")
        # Se quiser parar tudo ao primeiro erro, descomente a linha abaixo:
        # sys.exit(1)

def main():
    scripts = [
        r"C:\Users\MEUCOMPUTADOR\Desktop\projeto análise de empréstimo\baixar_anexos.exe",
        r"C:\Users\MEUCOMPUTADOR\Desktop\projeto análise de empréstimo\robo_analise.exe",
        r"C:\Users\MEUCOMPUTADOR\Desktop\projeto análise de empréstimo\enviar_analise.exe",
    ]

    interval = 5 * 60  # 5 minutos

    for idx, exe in enumerate(scripts, start=1):
        run_script(exe)
        if idx < len(scripts):
            logging.info(f"Aguardando {interval//60} minutos antes do próximo executável...")
            time.sleep(interval)

    logging.info("Execução de todos os scripts concluída com sucesso.")

if __name__ == "__main__":
    main()
