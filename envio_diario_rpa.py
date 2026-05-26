import win32com.client as win32
import pyautogui
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURAÇÕES ---
ARQUIVO_EXCEL = r"C:\Users\distr\Desktop\AUTOMAÇÕES\MAIO 2026\DIÁRIO MAI26\DIÁRIO DE VENDAS MAIO 2026.xlsx"
ABA_EQUIPES = "ACOMP DIÁRIO"
ABA_GESTAO = "GESTÃO"

chrome_user_data_dir = r"C:\Users\distr\Desktop\Perfil_Robo_WPP"
profile_dir = "Default"

# Configuração das Equipes (Mantenha os IDs dos grupos aqui)
equipes = {
    "RONALDO": {
        "id_grupo": "J4iaDNclmwK708RHi7elZp",
        "intervalo": "A1:S28",
        "linhas_ocultar": ["13:14"] 
    },
    "IVAN": {
        "id_grupo": "I9j342H2UZ1AaUfDc0XwCi",
        "intervalo": "A1:S38",
        "linhas_ocultar": ["13:28"]
    },
    "ADALBERTO": {
        "id_grupo": "BCNwGQbT53s5QTV8qZCk8i",
        "intervalo": "A1:S47",
        "linhas_ocultar": ["13:38"]
    },
    "FABIO": {
        "id_grupo": "Iw1ycaT43ab1JICRQaxwN3",
        "intervalo": "A1:S59",
        "linhas_ocultar": ["13:47"]
    },
    "BALCAO": {
        "id_grupo": "BqHQo56CbLWGfL8jDLqrXJ",
        "intervalo": "A1:S65",
        "linhas_ocultar": ["13:59"]
    },
    "CLEBSON": {
        "id_grupo": "GnFdjoef19N5lBB87esetG",
        "intervalo": "A1:S74",
        "linhas_ocultar": ["13:66"]
    },
    "ISAIAS": {
        "id_grupo": "IozGAhS9xwp2AdCCxRsdzq",
        "intervalo": "A1:S85",
        "linhas_ocultar": ["13:74"]
    }
}

# Configuração da Gerência : 558196011020 - Christopher / 558199939097 - Fábio
gerencia = {
    "telefone": "558196011020",
    "intervalo_gestao": "A1:K21"
}

# --- FUNÇÕES AUXILIARES ---
def abrir_chat_whatsapp(driver, identificador):
    # Se o identificador for só número, é mensagem privada. Se tiver letras, é grupo.
    if identificador.isdigit():
        url = f"https://web.whatsapp.com/send?phone={identificador}"
    else:
        url = f"https://web.whatsapp.com/accept?code={identificador}"
        
    if identificador not in driver.current_url:
        driver.get(url)
        time.sleep(8) # Aguarda o WhatsApp Web carregar a conversa
        
        # Se for grupo, tenta clicar em "Entrar" caso apareça
        if not identificador.isdigit():
            try:
                btn = driver.find_element(By.XPATH, "//div[contains(@role,'button') and (contains(., 'Entrar') or contains(., 'Join'))]")
                btn.click()
                time.sleep(3)
            except:
                pass

def achar_caixa_texto(driver):
    start = time.time()
    while time.time() - start < 15:
        try:
            # Procura a barra de digitação do WhatsApp
            box = driver.find_element(By.XPATH, "//div[@contenteditable='true' and @data-tab='10']")
            return box
        except:
            time.sleep(1)
    return None

def preparar_aba_equipe(ws, intervalo, linhas_ocultar):
    ws.Rows.Hidden = False # Exibe todas as linhas para resetar
    if linhas_ocultar:
        for linha in linhas_ocultar:
            ws.Rows(linha).Hidden = True # Oculta as linhas especificadas
    ws.Range(intervalo).CopyPicture(Format=2)
    time.sleep(2)

print("--- Iniciando Automação Mestre Híbrida ---")

os.system("taskkill /f /im excel.exe >nul 2>&1")
time.sleep(1)

try:
    print("\n1. Preparando o Chrome...")
    options = Options()
    options.add_argument(f"--user-data-dir={chrome_user_data_dir}")
    options.add_argument(f"--profile-directory={profile_dir}")
    prefs = {"protocol_handler.excluded_schemes": {"whatsapp": True}}
    options.add_experimental_option("prefs", prefs)
    
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=options)
    driver.maximize_window()
    driver.get("https://web.whatsapp.com")

    print("2. Abrindo o Excel...")
    excel = win32.DispatchEx('Excel.Application')
    excel.Visible = True 
    excel.DisplayAlerts = False 
    excel.AskToUpdateLinks = False 

    print("\n=======================================================")
    print("PAUSA DE SEGURANÇA:")
    print("1. Feche a tela de Ativação do Excel.")
    print("2. Confirme que o WhatsApp Web carregou.")
    print("3. Volte aqui e aperte ENTER.")
    print("=======================================================")
    input("Aperte ENTER para iniciar os envios: ")

    wb = excel.Workbooks.Open(ARQUIVO_EXCEL, ReadOnly=True, UpdateLinks=0)
    ws_equipes = wb.Sheets(ABA_EQUIPES)
    ws_gestao = wb.Sheets(ABA_GESTAO)

    driver.minimize_window()
    driver.maximize_window()
    time.sleep(1)

    # --- ETAPA 1: ENVIO INDIVIDUAL PARA AS EQUIPES ---
    print("\nIniciando envio para os grupos das equipes...")
    for equipe_key, dados in equipes.items():
        print(f"Processando equipe: {equipe_key}")
        abrir_chat_whatsapp(driver, dados["id_grupo"])
        
        caixa_texto = achar_caixa_texto(driver)
        if not caixa_texto:
            print(f"Erro: Caixa de texto não encontrada para {equipe_key}.")
            continue

        caixa_texto.click()
        pyautogui.write("Bom dia, segue o Acompanhamento Diario!")
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)

        preparar_aba_equipe(ws_equipes, dados["intervalo"], dados["linhas_ocultar"])
        
        caixa_texto.click()
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(3)
        pyautogui.write(f"*Resultados - {equipe_key}*")
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(3)

    # --- ETAPA 2: ENVIO GERAL PARA A GERÊNCIA (NO PRIVADO) ---
    print("\nIniciando envio direto para o Gerente...")
    abrir_chat_whatsapp(driver, gerencia["telefone"])
    caixa_texto = achar_caixa_texto(driver)
    
    if caixa_texto:
        caixa_texto.click()
        pyautogui.write("Bom dia, segue o acompanhamento diario geral")
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)

        # Copia e envia a aba de Gestão
        ws_gestao.Range(gerencia["intervalo_gestao"]).CopyPicture(Format=2)
        time.sleep(2)
        caixa_texto.click()
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(3)
        pyautogui.write("*Resumo geral - Gestão*")
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(3)

        # Copia e envia o detalhamento de cada equipe para o gerente
        for equipe_key, dados in equipes.items():
            preparar_aba_equipe(ws_equipes, dados["intervalo"], dados["linhas_ocultar"])
            caixa_texto.click()
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(3)
            pyautogui.write(f"*Detalhamento - {equipe_key}*")
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(3)

except Exception as e:
    print(f"Erro fatal: {e}")
finally:
    try:
        excel.CutCopyMode = False 
        wb.Close(SaveChanges=False)
        excel.Quit()
    except:
        pass

print("\nProcesso finalizado com sucesso.")