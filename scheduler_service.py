import time
from datetime import datetime
import auto_analise
import download_gmail
import sys

def run_scheduler():
    print("🕒 Serviço de Agendamento Iniciado")
    print("📅 Tarefas agendadas:")
    print("   - 06:30: Download de Arquivos (Gmail)")
    print("   - 07:00: Processamento e Atualização do Dashboard")
    print("---------------------------------------------------")
    
    # Estado da execução diária
    last_run_date = datetime.now().date()
    # Se iniciou depois do horário, assume que não fez ainda (ou deixa para o dia seguinte? Melhor deixar para dia seguinte para evitar duplicidade acidental, ou manual)
    # Resetando flags para o dia ATUAL
    download_done = False
    process_done = False
    
    while True:
        now = datetime.now()
        current_time = now.strftime("%H:%M")
        today_date = now.date()
        
        # Vira o dia
        if today_date != last_run_date:
            print(f"🌙 Virada de dia detectada: {last_run_date} -> {today_date}")
            last_run_date = today_date
            download_done = False
            process_done = False
        
        # Tarefa 1: Download (06:30)
        if current_time == "06:30" and not download_done:
            print(f"\n⏰ [06:30] Iniciando Download...")
            try:
                download_gmail.download_daily_attachments()
                print("✅ Download concluído (ou tentado).")
            except Exception as e:
                print(f"❌ Erro no download: {e}")
            
            download_done = True
            time.sleep(60) # Evita reexecução no mesmo minuto
            
        # Tarefa 2: Processamento (07:00)
        if current_time == "07:00" and not process_done:
            print(f"\n⏰ [07:00] Iniciando Análise Diária...")
            try:
                # Executa fluxo sem baixar (pois já baixou as 06:30)
                auto_analise.executar_fluxo_diario(baixar_email=False)
                print("✅ Processamento concluído.")
            except Exception as e:
                print(f"❌ Erro no processamento: {e}")
                
            process_done = True
            time.sleep(60)
            
        # Feedback de status (heartbeat a cada hora)
        if now.minute == 0 and now.second < 5:
            print(f"💓 Serviço ativo: {now.strftime('%d/%m %H:%M')}")
            time.sleep(5)
            
        time.sleep(10)

if __name__ == "__main__":
    try:
        run_scheduler()
    except KeyboardInterrupt:
        print("\n🛑 Serviço interrompido pelo usuário.")
        sys.exit(0)
