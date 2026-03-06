import pyautogui
import os
import time
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe
from config_local import CAMINHO_POWERBI, PASTA_DOWNLOADS, NOME_PLANILHA

# --- CONFIGURAÇÕES DE CAMINHOS ---
caminho_arquivo = CAMINHO_POWERBI
pasta_downloads = PASTA_DOWNLOADS
NOME_PLANILHA = NOME_PLANILHA

# --- CONFIGURAÇÕES GOOGLE SHEETS ---
NOME_PLANILHA = "Relatorio_Daily"
ABA_EXPORT = "Export_Oportunidades" 
ABA_EXPORT2 = "Export_Painel"

# --- COORDENADAS MAPEADAS ---
x_atu, y_atu = 338, 113
x_conf, y_conf = 344, 147
x_exp_menu, y_exp_menu = 754, 182
x_exp_opcao, y_exp_opcao = 842, 208
x_intermediario, y_intermediario = 804, 677
x_pos_final, y_pos_final = 720, 680

# Pausa padrão entre comandos para estabilidade
pyautogui.PAUSE = 1.0

# --- FUNÇÃO DE TEMPORIZADOR VISÍVEL ---
def aguardar_com_contagem(segundos, mensagem="Aguardando"):
    for i in range(segundos, 0, -1):
        print(f"   {mensagem}: {i}s restantes...   ", end="\r")
        time.sleep(1)
    print(f"\n   --- {mensagem} Finalizado ---")

# --- FUNÇÃO DE IMPORTAÇÃO CORRIGIDA ---
def importar_para_google_sheets(caminho_csv, nome_aba):
    try:
        print(f"\n--- Iniciando Integração: {nome_aba} ---")
        if not os.path.exists(caminho_csv):
            print(f"ERRO: Arquivo {caminho_csv} não encontrado.")
            return

        escopos = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        
        # --- AJUSTE DE CAMINHO AQUI ---
        # Pega a pasta onde está o script (src)
        diretorio_do_script = os.path.dirname(os.path.abspath(__file__))
        # Sobe um nível para a pasta raiz (onde está a pasta config)
        diretorio_raiz = os.path.dirname(diretorio_do_script)
        # Aponta para config/cred.json
        caminho_cred = os.path.join(diretorio_raiz, "config", "cred.json")
        
        if not os.path.exists(caminho_cred):
            print(f"ERRO CRÍTICO: cred.json não encontrado em {caminho_cred}")
            return

        credenciais = Credentials.from_service_account_file(caminho_cred, scopes=escopos)
        cliente = gspread.authorize(credenciais)
        
        planilha = cliente.open(NOME_PLANILHA)
        aba = planilha.worksheet(nome_aba)
        
        try:
            df = pd.read_csv(caminho_csv, encoding='utf-8', sep=',')
        except:
            df = pd.read_csv(caminho_csv, encoding='utf-16', sep='\t')
            
        df = df.fillna('')
        print(f"    Dados carregados: {len(df)} linhas.")

        print(f"    Limpando intervalo A2:AA999999...")
        aba.batch_clear(["A2:AA999999"])
        
        set_with_dataframe(aba, df, row=2, include_index=False, include_column_header=False)
        print(f"SUCESSO: Google Sheets atualizado ({nome_aba})!")

    except Exception as e:
        print(f"ERRO CRÍTICO na integração de {nome_aba}: {e}")

# --- INÍCIO DA EXECUÇÃO TOTAL ---

# PASSO 0: LIMPEZA
print(f"0. Limpando arquivos antigos em Downloads...")
for nome_arq in arquivos_para_limpar:
    caminho_full = os.path.join(pasta_downloads, nome_arq)
    if os.path.exists(caminho_full):
        try:
            os.remove(caminho_full)
            print(f"   - {nome_arq} excluído.")
        except: pass

# PASSO 1: ABERTURA POWER BI
print("\n1. Abrindo Power BI...")
os.startfile(caminho_arquivo)
aguardar_com_contagem(45, "Carregando Power BI")
pyautogui.click(500, 10) 
pyautogui.press('esc')
pyautogui.hotkey('alt', 'space')
pyautogui.press('x')
time.sleep(5)

# PASSO 2: ATUALIZAÇÃO
print("2. Iniciando atualização de dados...")
pyautogui.click(x=x_atu, y=y_atu)
time.sleep(3)
pyautogui.click(x=x_conf, y=y_conf)
aguardar_com_contagem(450, "Atualizando Dados")

# PASSO 3: EXPORTAÇÃO 1
print("3. Exportando Relatório 1...")
pyautogui.moveTo(x_exp_menu, y_exp_menu, duration=0.5)
pyautogui.mouseDown(); time.sleep(0.2); pyautogui.mouseUp()
time.sleep(3)
pyautogui.click(x_exp_opcao, y_exp_opcao)
aguardar_com_contagem(12, "Preparando CSV 1")
pyautogui.hotkey('ctrl', 'a'); pyautogui.press('backspace')
pyautogui.write(os.path.join(pasta_downloads, "Export"), interval=0.1)
pyautogui.press('enter')

# PASSO 4: INTERMEDIÁRIO
aguardar_com_contagem(15, "Trocando de aba")
pyautogui.click(x_intermediario, y_intermediario)

# PASSO 5: EXPORTAÇÃO 2
print("4. Exportando Relatório 2...")
time.sleep(5)
pyautogui.moveTo(x_exp_menu, y_exp_menu, duration=0.5)
pyautogui.mouseDown(); time.sleep(0.2); pyautogui.mouseUp()
time.sleep(3)
pyautogui.click(x_exp_opcao, y_exp_opcao)
aguardar_com_contagem(12, "Preparando CSV 2")
pyautogui.hotkey('ctrl', 'a'); pyautogui.press('backspace')
pyautogui.write(os.path.join(pasta_downloads, "Export2"), interval=0.1)
pyautogui.press('enter')

# PASSO 6: SALVAMENTO E FECHAMENTO
print(f"5. Finalizando Power BI...")
aguardar_com_contagem(15, "Salvando progresso")
pyautogui.click(x_pos_final, y_pos_final)
time.sleep(5)
pyautogui.hotkey('ctrl', 's') 
aguardar_com_contagem(15, "Fechando aplicativo")
pyautogui.hotkey('alt', 'f4')

# PASSO 7: INTEGRAÇÃO GOOGLE SHEETS
print("\n6. Enviando dados para o Google Sheets...")
time.sleep(5)
importar_para_google_sheets(os.path.join(pasta_downloads, "Export.csv"), ABA_EXPORT)
importar_para_google_sheets(os.path.join(pasta_downloads, "Export2.csv"), ABA_EXPORT2)

print("\n--- Automação Total Concluída com Sucesso ---")