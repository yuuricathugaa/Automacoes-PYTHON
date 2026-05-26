import win32com.client as win32
import pyautogui
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURAÇÕES ---
BASE_DIR = r"C:\Users\distr\Desktop\AUTOMAÇÕES\MAIO 2026\VALIDADAS INDUSTRIAS MAI26"
RANGE_PAINEL = "A1:H80" 

# Perfil do Robô (Garante que ele já entra logado no WhatsApp Web)
chrome_user_data_dir = r"C:\Users\distr\Desktop\Perfil_Robo_WPP"
profile_dir = "Default"

equipes = {
    "CLEBSON": {
        "arquivo": "INDÚSTRIAS CLEBSON MAI26.xlsx",
        "id_grupo": "GnFdjoef19N5lBB87esetG",
        "vendedores": ["GESTAO", "AMELIE", "MAURILIO", "DANILO", "DANIEL", "GIOVANI", "PAULO", "DIEGO"]
    },
    "BALCÃO": {
        "arquivo": "INDÚSTRIAS BALCÃO MAI26.xlsx",
        "id_grupo": "BqHQo56CbLWGfL8jDLqrXJ",
        "vendedores": ["GESTAO", "JOELMA", "ANDREA", "LORAINE", "LILLIAN", "VANILDO"]
    },
    "ISAIAS": {
        "arquivo": "INDÚSTRIAS ISAÍAS MAI26.xlsx",
        "id_grupo": "IozGAhS9xwp2AdCCxRsdzq",
        "vendedores": ["GESTAO", "JUNIOR", "ELAINE", "LUANA", "GENIVAL", "GILMAR", "ANDRE", "JHONNYS"]
    },
    "FÁBIO": {
        "arquivo": "INDÚSTRIAS FÁBIO MAI26.xlsx",
        "id_grupo": "Iw1ycaT43ab1JICRQaxwN3",
        "vendedores": ["GESTAO", "MARCONDES", "RENATA", "FAUSTO", "SAMUEL", "SIMONE", "DANIEL LIMA", "EDINALDO", "GIVALDO", "ORIOSVALDO", "MARIA EDUARDA", "PAULO"]
    },
    "ADALBERTO": {
        "arquivo": "INDÚSTRIAS ADALBERTO MAI26.xlsx",
        "id_grupo": "BCNwGQbT53s5QTV8qZCk8i",
        "vendedores": ["GESTAO", "PATRICIA ADRIANA", "RONILSON", "BRUNO", "GABRIEL CARLOS", "MARCIANO", "P HENRIQUE", "EVERTON", "IGO"]
    },
    "IVAN": {
        "arquivo": "INDÚSTRIAS IVAN MAI26.xlsx",
        "id_grupo": "I9j342H2UZ1AaUfDc0XwCi",
        "vendedores": ["GESTAO", "ALEXANDRE", "ALMIR", "M ADELINO", "M LIMA", "PATRICIA", "J FRANCISCO", "JACKSON", "GLAYDSON", "DEIVISON"]
    },
    "RONALDO": {
        "arquivo": "INDÚSTRIAS RONALDO MAI26.xlsx",
        "id_grupo": "J4iaDNclmwK708RHi7elZp",
        "vendedores": ["GESTAO", "GESTAO 1", "J ARAUJO", "J THIAGO", "DIEGO", "J FABIO", "VALTER", "GABRIEL", "MOACIR", "PEDRO", "J NILTON", "VANDAME", "SAVIO", "LUCELIA", "FLAVIO"]
    }
}

print("--- Iniciando Automação Mestre Híbrida ---")

# Faxina inicial para garantir que não há processos travados
os.system("taskkill /f /im excel.exe >nul 2>&1")
time.sleep(1)

try:
    print("\n🌐 1️⃣ Preparando o Chrome do Robô...")
    options = Options()
    options.add_argument(f"--user-data-dir={chrome_user_data_dir}")
    options.add_argument(f"--profile-directory={profile_dir}")
    
    # Bloqueia a mensagem de abrir o App Nativo
    prefs = {"protocol_handler.excluded_schemes": {"whatsapp": True}}
    options.add_experimental_option("prefs", prefs)
    
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=options)
    driver.maximize_window()
    
    # Já deixa o WhatsApp abrindo no fundo
    driver.get("https://web.whatsapp.com")

    print("\n📊 2️⃣ Abrindo o motor do Excel...")
    excel = win32.DispatchEx('Excel.Application')
    excel.Visible = True 
    excel.DisplayAlerts = False 
    excel.AskToUpdateLinks = False 

    print("\n=======================================================")
    print("🛑 PAUSA MANUAL DE SEGURANÇA:")
    print("1. Vá até o Excel que acabou de abrir e FECHE a tela de Ativação.")
    print("2. Confirme que o WhatsApp Web carregou no Chrome do Robô.")
    print("3. Volte aqui nesta tela preta e aperte ENTER.")
    print("=======================================================")
    input("👉 Aperte ENTER para soltar as amarras do robô: ")

    print("\n🚀 INICIANDO A MARATONA DE ENVIOS! (Não mexa no mouse)")

    for equipe_key, meta in equipes.items():
        arquivo_excel = os.path.join(BASE_DIR, meta["arquivo"])
        group_id = meta["id_grupo"]
        vendedores = meta["vendedores"]
        
        if not os.path.exists(arquivo_excel):
            print(f"⚠️ Arquivo não encontrado: {arquivo_excel}. Pulando equipe...")
            continue
            
        print(f"\n📂 Carregando a planilha da equipe {equipe_key}...")
        wb = excel.Workbooks.Open(arquivo_excel, ReadOnly=True, UpdateLinks=0)

        print(f"🔗 O Selenium está entrando no grupo da equipe {equipe_key}...")
        driver.get(f"https://web.whatsapp.com/accept?code={group_id}")
        time.sleep(8) # Tempo generoso para carregar a conversa

        # Tenta clicar no botão "Entrar na conversa" caso apareça
        try:
            btn = driver.find_element(By.XPATH, "//div[contains(@role,'button') and (contains(., 'Entrar') or contains(., 'Join'))]")
            btn.click()
            time.sleep(3)
        except:
            pass

        # Força o Chrome a vir para frente no Windows para o PyAutoGUI funcionar
        driver.minimize_window()
        driver.maximize_window()
        time.sleep(1)

        # O Selenium acha a caixa de texto e clica nela
        try:
            main_box = driver.find_element(By.XPATH, "//div[@contenteditable='true' and @data-tab='10']")
            main_box.click()
            time.sleep(1)
        except Exception as e:
            print(f"❌ Não achei a caixa de texto do grupo {equipe_key}. Pulando equipe.")
            wb.Close(SaveChanges=False)
            continue

        # PyAutoGUI assume o teclado e envia o texto inicial
        pyautogui.write("Segue a Parcial de Industrias!")
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)

        # Começa os recortes
        for vendedor in vendedores:
            try:
                print(f"📸 Copiando {vendedor}...")
                ws = wb.Sheets(vendedor)
                ws.Range(RANGE_PAINEL).CopyPicture(Format=2) 
                time.sleep(2)
                
                # Clica na caixa de texto do Chrome pra garantir o foco absoluto
                main_box.click()
                time.sleep(0.5)
                
                # PyAutoGUI cola e envia!
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(3) 
                
                pyautogui.write(f"*{vendedor}*")
                time.sleep(1)
                pyautogui.press('enter')
                
                print(f"✅ Enviado: {vendedor}")
                time.sleep(4) 
                
            except Exception as e:
                print(f"❌ Erro ao enviar {vendedor}: {e}")

        # Fecha a planilha da equipe atual e vai para a próxima
        wb.Close(SaveChanges=False)
        print(f"✅ Equipe {equipe_key} finalizada com sucesso!")

except Exception as erro_geral:
    print(f"❌ Erro fatal: {erro_geral}")
finally:
    try:
        excel.CutCopyMode = False 
        excel.Quit()
        driver.quit()
    except:
        pass

print("\n🏁 Processo de todas as equipes finalizado com sucesso!")