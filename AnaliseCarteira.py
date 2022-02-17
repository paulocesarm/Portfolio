### IMPORTANDO BIBLIOTECAS
from typing import Concatenate
import pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt
import pandas_datareader.data as web
import win32com.client as win32
from telepot import Bot
import time

### CAMINHOS DE ARQUIVO
CAMINHO_CARTEIRA = Path('C:/Users/paulo/Desktop/paulocesarm/python study/#HASHTAG/ENCONTROS AO VIVO/03 - ANÁLISES FINANCEIRS/CarteiraMentoria.xlsx')
CAMINHO_TD = 'https://www.tesourotransparente.gov.br/ckan/dataset/f0468ecc-ae97-4287-89c2-6d8139fb4343/resource/e5f90e3a-8f8d-4895-9c56-4bb2f7877920/download/VendasTesouroDireto.csv'
tesouro_direto = 'https://www.tesourotransparente.gov.br/ckan/dataset/f0468ecc-ae97-4287-89c2-6d8139fb4343/resource/e5f90e3a-8f8d-4895-9c56-4bb2f7877920/download/VendasTesouroDireto.csv'

### CAMINHO PARA SALVAR GRÁFICOS
VALOR_INVESTIDO_ATIVO = r'C:/Users/paulo/Desktop/paulocesarm/python study/#HASHTAG/ENCONTROS AO VIVO/03 - ANÁLISES FINANCEIRS/valor_investido_ativo.png'
VALOR_INVESTIDO_TIPO = r'C:/Users/paulo/Desktop/paulocesarm/python study/#HASHTAG/ENCONTROS AO VIVO/03 - ANÁLISES FINANCEIRS/valor_investido_tipo.png'
VALOR_IBOV = r'C:/Users/paulo/Desktop/paulocesarm/python study/#HASHTAG/ENCONTROS AO VIVO/03 - ANÁLISES FINANCEIRS/valores_ibov.png'
COMPARATIVO_CARTEIRA_IBOV = r'C:/Users/paulo/Desktop/paulocesarm/python study/#HASHTAG/ENCONTROS AO VIVO/03 - ANÁLISES FINANCEIRS/comparativo_carteira_ibov.png'

### PARÂMETROS DO BOT TELEGRAM
TOKEN = '5085552159:AAGRv9GTciGveMd6aZYBlC4NSwbCmgjIHrE' #token do bot
SeuID = 5026573457 #id do usuario
id_grupo = -629162419 #id do grupo

def enviando_telegram():
    bot = Bot(TOKEN) #inicializando bot
    bot.sendMessage(SeuID,"Gráfico do valor investido por ativo: ")
    bot.sendPhoto(chat_id=SeuID, photo=open(VALOR_INVESTIDO_ATIVO, 'rb'))
    time.sleep(2)
    bot.sendMessage(SeuID,"Gráfico do valor investido por tipo: ")
    bot.sendPhoto(chat_id=SeuID, photo=open(VALOR_INVESTIDO_ATIVO, 'rb'))
    time.sleep(2)
    bot.sendMessage(SeuID,"Gráfico do Ibovespa pelo tempo: ")
    bot.sendPhoto(chat_id=SeuID, photo=open(VALOR_IBOV, 'rb'))
    time.sleep(2)
    bot.sendMessage(SeuID,"Gráfico Carteira x IBOV: ")
    bot.sendPhoto(chat_id=SeuID, photo=open(COMPARATIVO_CARTEIRA_IBOV, 'rb'))
    return print("Mensagem enviada no Telegram")

def enviando_email():

    outlook = win32.Dispatch('outlook.application') #definindo aplicação pelo outlook
    mail = outlook.CreateItem(0) #criando o email
    mail.To = 'paulocesaradvice@gmail.com' #defindo o destinatário
    mail.Subject = 'Relatório Gerencial da Carteira' #definindo título
    mail.Body = f''' Olá Paulo, tudo bem?
    Conforme solicitado, levantamos os principais gráficos referente a sua carteira e enviamos em anexo.
    Caso tenha alguma dúvida, não hesite em entrar em contato conosco.''' #corpo do email#corpo do email
    mail.Attachments.Add(VALOR_INVESTIDO_ATIVO) #adicionando anexo
    mail.Attachments.Add(VALOR_INVESTIDO_TIPO) #adicionando anexo
    mail.Attachments.Add(VALOR_IBOV) #adicionando anexo
    mail.Attachments.Add(COMPARATIVO_CARTEIRA_IBOV) #adicionando anexo
    mail.Send() #enviando :)

    return print("Email Enviado")


def main():

    ### PEGANDO COTAÇÕES
    ## indices: ^iINDICE
    ## ações brasileiras: TICKER.SA
    ## source = yahoo
    ## datas: 'aaaa-mm-dd'

    ## IBOVESPA
    ibov_df = web.DataReader('^BVSP', data_source='yahoo', start = '2020-01-02', end = '2020-12-10') #web.DataReader('TICKER', data_source='fonte', start = 'data inicio', end = 'data final')
    print(ibov_df.info()) #printando info do dataframe para análise

    ## MINHA CARTEIRA
    carteira = pd.read_excel(CAMINHO_CARTEIRA)
    carteira_tipo = carteira.groupby('Tipo')['Valor Investido'].sum()
    carteira_df = pd.DataFrame()
    for ativo in carteira['Ativos']:
        if 'Tesouro' not in ativo: #retirando o tesouro direto do FOR (Não da para retirar as informações dele pelo link do yahoo)
            carteira_df[ativo] = web.DataReader('{}.SA'.format(ativo), data_source='yahoo', start = '2020-01-01', end = '2020-12-10')['Adj Close'] #pegando os dados dos ativos da carteira
        else:
            pass

    carteira_df = carteira_df.ffill() #preenche o vazio com o dado da linha anterior, a coluna do SMALL11 estava faltando 2 dados. Fizemos isso para preencher.
    print(carteira_df.info()) #printando info do dataframe para análise

    ## TESOURO
    tesouro_df = pd.read_csv(CAMINHO_TD, sep = ';', decimal = ',') #lendo o arquivo CSV com os dados do tesouro (através de um link direto de csv)
    print(tesouro_df.info()) #printando informações do dataframe para análise
    tesouro_df['Data Venda'] = pd.to_datetime(tesouro_df['Data Venda'], format = '%d/%m/%Y') #passando o formato de data de dd/mm/aaaa para o datetime padrão do pandas ('aaaa-mm-dd'), o format não é o formato que ele vai ficar e sim o que ele é.
    tesouro_df = tesouro_df[tesouro_df['Tipo Titulo'] == 'Tesouro Selic'] #pegando somente as linhas que o tipo título é o tesouro selic
    tesouro_df = tesouro_df[['Data Venda', 'PU']] #filtrando base para pegar apenas as colunas de Data Venda e PU
    tesouro_df = tesouro_df.rename(columns = {'Data Venda':'Date' , 'PU': 'Tesouro Selic'}) #renomeando as colunas de Data Venda e PU.

    ## UNINDO DADOS DA MINHA CATEIRA COM O TESOURO
    carteira_df = carteira_df.merge(tesouro_df[['Date','Tesouro Selic']], on = 'Date', how = 'left') #colando os dados do lado e juntando a coluna em comum. (sempre tem que ter uma coluna em comum para usar o merge!)
    print(carteira_df.info()) #printando informações do dataframe para análise

    ### REALIZANDO CALCULO DE INDICADORES
    ## TRATANDO DADOS
    carteira_copia = carteira_df.copy() #criando uma cópia do dataframe da carteira, para não mudar o original na hora dos cálculos

    for i,ativo in enumerate(carteira['Ativos']):
        carteira_copia[ativo] = carteira_copia[ativo]*carteira.loc[i, 'Qtde'] #substituindo os valores das colunas dos ativos pelo Total Investido

    carteira_copia = carteira_copia.set_index('Date') #setando a coluna de data para ser o indice para que no SUM ela não seja contabilizada.

    carteira_copia['Total'] = carteira_copia.sum(axis = 1) #criando coluna de total somando as linhas das colunas restantes (o axis = 1 muda o padrão de somar coluna para somar linha)
    ## CALCULANDO INDICADORES
    carteira_copia_norm = carteira_copia/carteira_copia.iloc[0] #colocando na mesma base para que a comparação no gráfico seja real (partam do mesmo ínicio)
    ibov_df_norm = ibov_df/ibov_df.iloc[0] #colocando na mesma base para que a comparação no gráfico seja real
    rentabilidde_carteira = carteira_copia['Total'].iloc[-1]/carteira_copia['Total'].iloc[0] - 1 #calculando a rentabilidade da carteira
    rentabilidade_ibov = ibov_df['Adj Close'].iloc[-1]/ibov_df['Adj Close'].iloc[0] - 1 #calculando a rentabilidade do ibovespa
    print('Rentabilidade da Carteira {:.1%}'.format(rentabilidde_carteira)) #printando a rentabilidade da carteira
    print('Rentabilidade da IBOV {:.1%}'.format(rentabilidade_ibov)) #printando a rentabilidade do ibovespa

    # ### PLOTANDO GRÁFICOS
    ## PLOTANDO GRÁFICO COMPARATIVO DA CATEIRA COM O IBOVESPA
    grafico4 = carteira_copia_norm['Total'].plot(figsize = (15,5), label = 'Carteira')
    grafico4 = ibov_df_norm['Adj Close'].plot(figsize = (15,5), label = 'IBOV')
    plt.legend()
    plt.show()
    fig4 = grafico4.figure
    fig4.savefig(COMPARATIVO_CARTEIRA_IBOV)
    ## PLOTANDO GRÁFICO COM OS VALORES DO IBOV AO LONGO DO TEMPO
    grafico3 = ibov_df['Adj Close'].plot(figsize = (15,5))
    grafico3.set_xlabel('')
    fig3 = grafico3.figure #atribuindo gráfico a essa variavel
    fig3.savefig(VALOR_IBOV) #salvando gráfico no caminho especificado
    ## PLOTANDO GRÁFICO DE DISTRIBUIÇÃO DO VALOR INVESTIDO EM PERCENTUAL POR ATIVO DA CARTEIRA
    grafico1 = carteira['Valor Investido'].plot.pie(labels = carteira['Ativos'], y = 'Valor Investido', legend = False, title ="Distruibuição dos ativos da Carteira",figsize = (15,5), autopct = "%.1f%%") # plotando gráfico de pizza
    grafico1.set_ylabel('') #tirando título do eixo y
    grafico1.set_xlabel('') #tirando título do eixo x
    fig1 = grafico1.figure #atribuindo gráfico a essa variavel
    fig1.savefig(VALOR_INVESTIDO_ATIVO) #salvando gráfico no caminho especificado
    ## PLOTANDO GRÁFICO DE DISTRIBUIÇÃO DO VALOR INVESTIDO EM PERCENTUAL POR TIPO DE ATIVO DA CARTEIRA
    grafico2 = carteira.groupby('Tipo')['Valor Investido'].sum().plot.pie(y = 'Valor Investido', legend = False, title ="Distruibuição dos tipos da Carteira",figsize = (15,5), autopct = "%.1f%%")
    grafico2.set_ylabel('') #tirando o título do eixoo y do gráfico 2
    grafico2.set_xlabel('') #tirando título do eixo x grafico 2
    fig2 = grafico2.figure #atribuindo gráfico a essa variavel
    fig2.savefig(VALOR_INVESTIDO_TIPO)

    return None

print("Iniciando rotina de análise da carteira")
função = main()
print("Iniciando Rotina de envio para o email")
enviar_email = enviando_email()
print("Iniciando Rotina de envio para o telegram")
telegram = enviando_telegram()




