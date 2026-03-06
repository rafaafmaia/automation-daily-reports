import pyautogui
import time

# Aviso inicial
pyautogui.alert("Ao clicar em OK, você terá 5 segundos para posicionar o mouse sobre o botão ATUALIZAR.")

# Contagem regressiva simples
time.sleep(5)

# Pega a posição
x, y = pyautogui.position()

# Exibe o resultado em uma janela na sua frente
mensagem = f"COORDENADAS CAPTURADAS!\n\nX = {x}\nY = {y}\n\nAnote esses números e feche esta janela."
pyautogui.alert(mensagem)

print(f"Para caso de cópia: X={x}, Y={y}")