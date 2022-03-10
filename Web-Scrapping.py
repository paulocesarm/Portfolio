###IMPORTANDO BIBLIOTECAS
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import time
import zipfile
from pathlib import Path
import pandas as pd
import win32com.client as win32

###CAMINHOS
CAMINHO_ZIP = Path(r'C:/Users/paulo/Downloads/archive.zip')
CAMINHO_ARQUIVO_DESTINO = Path(r'C:/Users/paulo/Desktop/paulocesarm/python study/#HASHTAG/ENCONTROS AO VIVO/02 - AUTOMAÇÃO/Arquivos')
CAMINHO_ARQUIVO_BANCO = r'C:/Users/paulo/Desktop/paulocesarm/python study/#HASHTAG/ENCONTROS AO VIVO/02 - AUTOMAÇÃO/Arquivos/BankChurners.csv'

###URL SITE
URL = 'https://www.kaggle.com/sakshigoyal7/credit-card-customers'

###DEFININDO FUNÇÃO DE WEBSCRAPPING
def webscrapping():
    driver.get(URL)
    driver.find_element_by_xpath('//*[@id="site-content"]/div[3]/div[1]/div/div/div[2]/div[2]/div[1]/div[2]/a/button/span').click()
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="site-container"]/div[1]/div/form/div[2]/div/div[2]/a/li/div/span').click()
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="site-container"]/div[1]/div/form/div[2]/div[1]/div/label').send_keys('paulocesarmlf@gmail.com')
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="site-container"]/div[1]/div/form/div[2]/div[2]/div/label/input').send_keys('sfclove100')
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="site-container"]/div[1]/div/form/div[2]/div[3]/button/span').click()
    time.sleep(3)
    driver.find_element_by_xpath('//*[@id="site-content"]/div[3]/div[1]/div/div/div[2]/div[2]/div[1]/div[2]/a/button/span').click()
    time.sleep(10)
    driver.close()
    return print("Base baixada do kaggle")

###DESCOMPACTANDO ARQUIVO ZIP
def descompact_zip():
    with zipfile.ZipFile(CAMINHO_ZIP,'r') as zip_ref:
        zip_ref.extractall(CAMINHO_ARQUIVO_DESTINO)
    return print("Arquivo descompactado")

time.sleep(60)

###DEFININDO FUNÇÃO QUE ENVIA O EMAIL
def enviando_email(status,cartao,tempomedio):

    outlook = win32.Dispatch('outlook.application') #definindo aplicação pelo outlook
    mail = outlook.CreateItem(0) #criando o email
    mail.To = 'paulocesaradvice@gmail.com' #defindo o destinatário
    mail.Subject = 'Relatório' #definindo título
    mail.Body = f''' Olá Paulo, tudo bem?
    Conforme solicitado, levantamos os principais indicadores dos nossos clientes.
    Temos atualmente a seguinte divisão da base de clientes:

    {status}

    Quando analisamos os clientes ativos, percebemos a seguinte divisão de categorias de cartão:

    {cartao}

    Já quanto ao tempo médio de permanência dos clientes temos em média {tempomedio} meses ''' #corpo do email

    mail.Attachments.Add(CAMINHO_ARQUIVO_BANCO) #adicionando anexo
    mail.Send() #enviando :)

    return print("Email Enviado")

###TRATANDO OS DADOS E CALCULANDO INDICADORES
def main():
    ##TRATANDO DADOS
    banco_df = pd.read_csv(CAMINHO_ARQUIVO_BANCO, sep = ',') #lendo o arquivo 
    resumo_status = banco_df.groupby('Attrition_Flag')['Attrition_Flag'].count() #Agrupando pela coluna Attrition_Flag e contando a mesma coluna
    banco_df_filtrado = banco_df[banco_df['Attrition_Flag'] == 'Existing Customer'] #filtrando o dataframe onde banco_df['Attrition_Flag'] == 'Existing Customer' (clientes ativo)
    resumo_cartao = banco_df_filtrado.groupby('Card_Category')['Card_Category'].count() #Agrupando pela coluna Card_Category e contando a mesma coluna
    resumo_cartao.index.names = ['Categoria do Card - Existing Customers'] #mudando o nome do index desse card
    ##CALCULANDO INDICADORES
    tempo_medio = banco_df['Months_on_book'].mean() #calculando a media de meses (geral) no book
    limite_todomundo = banco_df['Credit_Limit'].mean() #calculando a media do limite (geral) de crédito

    banco_df_filtrado2 = banco_df[banco_df['Attrition_Flag'] == 'Attrited Customer'] #filtrando dataframe para ex-clientes
    limite_excliente = banco_df_filtrado2['Credit_Limit'].mean() #calculando média de limite de crédito apenas para os ex-clientes
    print("Se preparando para enviar email...")
    enviar_email = enviando_email(resumo_status.to_string(),resumo_cartao.to_string(),tempo_medio)

    return None

driver= webdriver.Chrome(ChromeDriverManager().install()) #instalando drive do chrome
driver.maximize_window() #maximando janela
coletar_dados = webscrapping()
time.sleep(10) #esperando 10 segundos
descompactando_arquivo = descompact_zip()
time.sleep(60) #esperando 60 segundos
print('Rodando Função Principal')
principal = main()
