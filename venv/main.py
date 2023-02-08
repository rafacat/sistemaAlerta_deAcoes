#Importando bibliotecas

#import pandas as pd
import datetime
import yfinance as yf
from matplotlib import pyplot as plt
import mplcyberpunk
import win32com.client as winzinho

#Entra com as carteiras que gostaria de saber a cotação
siglas_deNegociacoes = ['GC=F','^BVSP','TSLA']

#Mostra para o sistema como descobrir o alcance anual
hoje = datetime.datetime.now()
um_anoAtras = hoje - datetime.timedelta(days = 365)

#Busca as informações financeiras
dados_do_mercado = yf.download(siglas_deNegociacoes, um_anoAtras ,hoje)
###print(dados_do_mercado)

#Filtra fechamento ajustado
dados_fechamento = dados_do_mercado['Adj Close']

#Nomeia colunas
dados_fechamento.columns = ['gold Apr 23', 'ibovespa', 'tesla']

#Remove dados em branco
dados_fechamento = dados_fechamento.dropna()
###print(dados_fechamento)

#Informa o fechamento anual e mensal das ações
dados_anuais = dados_fechamento.resample("Y").last()
dados_mensais = dados_fechamento.resample("M").last()

#calcula o fechamento do dia e retorna o fechamento diario e mensal
retorno_anual = dados_anuais.pct_change().dropna()
retorno_mensal = dados_mensais.pct_change().dropna()
retorno_diario = dados_fechamento.pct_change().dropna()

#Retorna ultimo valor diario, mensal e anual
retorno_gold_diario = retorno_diario.iloc[-1,0]
retorno_ibov_diario = retorno_diario.iloc[-1,1]
retorno_tesla_diario = retorno_diario.iloc[-1,2]

retorno_gold_mensal = retorno_mensal.iloc[-1,0]
retorno_ibov_mensal = retorno_mensal.iloc[-1,1]
retorno_tesla_mensal = retorno_mensal.iloc[-1,2]

retorno_gold_anual = retorno_anual.iloc[-1,0]
retorno_ibov_anual = retorno_anual.iloc[-1,1]
retorno_tesla_anual = retorno_anual.iloc[-1,2]

#Formatando casas decimais
retorno_gold_diario = round((retorno_gold_diario * 100), 2)
retorno_ibov_diario = round((retorno_ibov_diario * 100), 2)
retorno_tesla_diario = round((retorno_tesla_diario * 100), 2)

retorno_gold_mensal = round((retorno_gold_mensal * 100), 2)
retorno_ibov_mensal = round((retorno_ibov_mensal * 100), 2)
retorno_tesla_mensal = round((retorno_tesla_mensal * 100), 2)

retorno_gold_anual = round((retorno_gold_anual * 100), 2)
retorno_ibov_anual = round((retorno_ibov_anual * 100), 2)
retorno_tesla_anual = round((retorno_tesla_anual * 100), 2)

#Cria um grafico comparativo das as ações
plt.style.use("cyberpunk")

dados_fechamento.plot(y = "gold Apr 23", use_index = True, legend = False)
plt.title("Gold Apr 23")
#plt.savefig('gold.png', dpi = 300)

dados_fechamento.plot(y = "ibovespa", use_index = True, legend = False)
plt.title("Ibovespa")
#plt.savefig('ibov.png', dpi = 300)

dados_fechamento.plot(y = "tesla", use_index = True, legend = False)
plt.title("tesla")
#plt.savefig('tesla.png', dpi = 300)

outlook = winzinho.Dispatch('Outlook.Application')
email = outlook.CreateItem(0)

email.To = "rafael-spi@hotmail.com"
email.Subject = "Teste"
email.Body = f''' Olá tudo bem ??? Segue o relatorio diario:

Bolsa:

No ano o "Ibovespa" esta tendo uma rentabilidade de {retorno_ibov_anual}%,
enquanto no mês a rentabilidade é de {retorno_ibov_mensal}%.

Gold:

O ouro esta tendo uma rentabilidade de {retorno_gold_anual}%,
anualmente, enquanto no mês a rentabilidade de {retorno_gold_mensal}%.

Tesla:

Atualmente as ações da "Tesla" esta dando um retorno anual de {retorno_tesla_anual}%,
e {retorno_tesla_mensal}% mensalmente.

Att,

Rafael Simião Pereira
CEO (TWc)

'''

anexo_ibovespa = r'C:\workspace\dev\Finanças\ibovespa_Dolar\ibov.png'
anexo_gold = r'C:\workspace\dev\Finanças\ibovespa_Dolar\gold.png'
anexo_tesla = r'C:\workspace\dev\Finanças\ibovespa_Dolar\tesla.png'

email.Attachments.Add(anexo_ibovespa)
email.Attachments.Add(anexo_gold)
email.Attachments.Add(anexo_tesla)

email.Send()
