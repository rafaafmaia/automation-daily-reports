import pyautogui
import os
import time
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe

# --- SISTEMA DE SEGURANÇA (IMPORTAÇÃO PROTEGIDA) ---
# Tenta carregar as configurações do seu arquivo local que está no .gitignore.
# Se o arquivo não existir (como no GitHub), ele usa nomes genéricos.
try:
    from config_local import CAMINHO_POWERBI, PASTA_DOWNLOADS, NOME_PLANILHA
except ImportError:
    # Estes valores aparecem no GitHub para proteger os dados reais da empresa
    CAMINHO_POWERBI = "C:/Caminho/Exemplo/Seu_Relatorio.pbix"
    PASTA_DOWNLOADS = "C:/Users/Usuario/Downloads"
    NOME_PLANILHA = "NOME_DA_SUA_PLANILHA_GOOGLE"

# Atribuição das variáveis de caminho
caminho_arquivo = CAMINHO_POWERBI
pasta_downloads = PASTA_DOWNLOADS
nome_planilha_google = NOME_PLANILHA

# --- CONFIGURAÇÕES DO GOOGLE SHEETS ---
NOME_PLANILHA_INTERNA = "Relatorio_Daily"
ABA_EXPORT = "Export_Oportunidades"
ABA_EXPORT2 = "Export_Painel"

# --- COORDENADAS MAPEADAS (Mantenha as suas originais aqui) ---
x_atu, y_atu = 338, 113
x_conf, y_conf = 344, 147
x_exp_menu, y_exp_menu = 754, 182

def iniciar_automacao():
    """
    Executa a rotina de abertura do Power BI, interação com a interface
    e exportação dos dados para o Google Sheets.
    """
    print(f"Iniciando automação para o arquivo: {caminho_arquivo}")
    
    # 1. Abrir o arquivo do Power BI
    if os.path.exists(caminho_arquivo):
        os.startfile(caminho_arquivo)
        time.sleep(30)  # Aguarda o software carregar
    else:
        print("ERRO: O caminho do Power BI não foi encontrado localmente.")
        return

    # 2. Exemplo de interação com PyAutoGUI (Sua lógica de cliques)
    # pyautogui.click(x_atu, y_atu)
    # time.sleep(5)
    
    print("Automação de interface concluída. Iniciando integração com Google Sheets...")

    # 3. Integração com Google Sheets (Usando credenciais protegidas pelo .gitignore)
    try:
        # O arquivo config/cred.json NÃO sobe para o GitHub
        scope = ['https://www.googleapis.com/auth/spreadsheets',
                 'https://www.googleapis.com/auth/drive']
        
        creds = Credentials.from_service_account_file('config/cred.json', scopes=scope)
        client = gspread.authorize(creds)
        
        # Abre a planilha pelo nome configurado
        sheet = client.open(nome_planilha_google).worksheet(ABA_EXPORT)
        
        # Exemplo de leitura de arquivo exportado pelo Power BI
        # df = pd.read_csv(os.path.join(pasta_downloads, "data.csv"))
        # set_with_dataframe(sheet, df)
        
        print("Dados enviados com sucesso!")
        
    except FileNotFoundError:
        print("Erro: O arquivo config/cred.json não foi encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro na integração: {e}")

if __name__ == "__main__":
    iniciar_automacao()