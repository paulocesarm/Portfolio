import pyautogui
import time
import pandas as pd
from pathlib import Path
import pyperclip
# Documentação https://pyautogui.readthedocs.io/en/latest/quickstart.html
# pyautogui.write() > Escreve
# pyautogui.click() > Clica
# pyautogui.locateOnScreen > Identifica as coordenadas da imagem que você passou como parametro
#pyautogui.locateAllOnScreen > Identifica as coordenadas de todas as imagens iguais da qual você passou como parametro, a resposta pode ser tratada como uma lista.

# pyautogui.hotkey > Usa atalhos do teclado (combinação de teclas)
# pyautogui.press > aperta um botão do teclado

pyautogui.alert("O código vai começar, não mexa em nada enquanto o código estiver rodando. Quando ele terminar eu te aviso")

pyautogui.PAUSE = 1 #esse funciona que nem o time mas para todos os comandos do pyautogui
pyautogui.press('win') #apertando tecla windwos
pyautogui.write('brave') #digitando brave na aba de pesquisa   
pyautogui.press('enter') #apertando enter para inicializar
pyautogui.hotkey('alt','space','x') #maximizando a tela


pyautogui.write('gmail') #escrevendo gmail no navegador
pyautogui.press('enter') #dando enter para ir para o que foi escrito no navegador
##É preciso tirar o print da local que você quer fazer algo e passar como parametro o caminho de onde foi salvo esse print
x,y, largura, altura = pyautogui.locateOnScreen('busca_google.png') # descobrindo posicao x, posicao y, largura e altura do local da imagem localizada.

pyautogui.click(x + largura/2, y + altura/2) # clicando na posição central da imagem localizada já que esse artificio (x+l/2, y+A/2) encontra exatamente o centro da imagem.
time.sleep(10) # se o codigo for muito grande e não queira esperar 10s use While not pyautogui.locateOnScree('teste_gmail') >> Time.sleep(1)

xmenu,ymenu,larguramenu,alturamenu = pyautogui.locateOnScreen('menu.png') #localizando a imagem do menu
pyautogui.click(xmenu + larguramenu/2, ymenu + alturamenu/2) #clicando no centro do menu

xcontato, ycontato, larguracontato, alturacontato = pyautogui.locateOnScreen('contatos.png') #localizando a imagem de contatos
pyautogui.click(xcontato + larguracontato/2, ycontato + alturacontato/2) #clicando no centro de contatos
time.sleep(1)
xexportar,yexportar,larguraexportar,alturaexportar = pyautogui.locateOnScreen('exportar.png') #localizando a imagem de exportar
pyautogui.click(xexportar + larguraexportar/2, yexportar + alturaexportar/2) #clicando no de exportar

xconfirmar,yconfirmar,larguraconfirmar,alturaconfirmar = pyautogui.locateOnScreen('confirmar_exportar.png') #localizando a imagem de confirmação do exportar
pyautogui.click(xconfirmar + larguraconfirmar/2, yconfirmar + alturaconfirmar/2) #clicando no centro da confirmação de exportar
time.sleep(1)
# xdown,ydown,larguradown,alturadown = pyautogui.locateOnScreen('download.png') #localizando a imagem de confirmação de download
# pyautogui.click(xdown + larguradown/2, ydown + alturadown/2) #clicando no centro de download
pyautogui.press('enter') #apertando enter para salvar o arquivo


df = pd.read_csv(r'C:/Users/paulo/Downloads/contacts.csv') #lendo o csv
df = df.dropna(axis = 1) #exclui vazios, como o axis = 1 ta excluindo colunas

pyautogui.hotkey('ctrl', 'pgup') # voltando para a pagina geral do gmail, o hotket faz uma combinação das teclas
xescrever,yescrever,larguraescrever,alturaescrever = pyautogui.locateOnScreen('escrever.png') #localizando a imagem de confirmação de escrever
pyautogui.click(xescrever + larguraescrever/2, yescrever + alturaescrever/2) #clicando no centro da escrever

pyautogui.write('paulocesaradvice@gmail.com') # escrevendo o email do destinatario
pyautogui.press('enter') #confirmando o email escrito
pyautogui.press('tab') # passando para o assunto do email
pyautogui.write('Teste Pyautogui') # escrevendo o assunto do email
pyautogui.press('tab') # passando para o texto a ser escrito.

texto = '''
Fala pauleta, ta ficando o mago no python ein....
Até no pyautogui paizera, o job na amazon ta bem ai einnnn

'''
pyperclip.copy(texto) # copiando o texto pq se digitar o texto direto lá ele vai considerar o teclado americano
pyautogui.hotkey('ctrl','v') #copiando o texto no canto de escrever
xenviar,yenviar,larguraenviar,alturaenviar = pyautogui.locateOnScreen('enviar.png') # localizando o botão de enviar
pyautogui.click(xenviar + larguraenviar/2,yenviar + alturaenviar/2) # enviando

pyautogui.alert("O código acabou")


###COMO DESCOBRI POSIÇÃO DE ONDE O MOUSE ESTÁ SEM SER PELA IMAGEM
# print(pyautogui.position()), vai te passar a posição exata de onde o seu mouse está.


