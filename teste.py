import pyautogui

# O caminho para a imagem que você deseja localizar
# Certifique-se de que 'image_0be63e.png' está na mesma pasta que este script,
# ou forneça o caminho completo para a imagem.
imagem_para_localizar = 'image.png'

print(f"Procurando a imagem: {imagem_para_localizar} na tela...")

try:
    localizacao = pyautogui.locateOnScreen(imagem_para_localizar, confidence=0.9)

    if localizacao:
        print(f"Imagem encontrada! Coordenadas (esquerda, topo, largura, altura): {localizacao}")

        # Se quiser clicar no centro da imagem:
        centro_x, centro_y = pyautogui.center(localizacao)
        print(f"Centro da imagem: X={centro_x}, Y={centro_y}")

        pyautogui.moveTo(centro_x, centro_y, duration=0.5) # Move o mouse lentamente
        pyautogui.click() # Clica no centro da imagem
        print("Mouse movido e clicado na imagem.")

    else:
        print("Imagem NÃO encontrada na tela.")

except pyautogui.ImageNotFoundException:
    print("Erro: A imagem não foi encontrada na tela. Certifique-se de que a imagem esteja visível.")
except Exception as e:
    print(f"Ocorreu um erro: {e}")

print("Fim do script.")