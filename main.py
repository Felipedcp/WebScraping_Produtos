from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from funcoes_de_busca import buscar_ofertas_produtos, enviar_ofertas_email
import pandas as pd

# criar um navegador
servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=servico)

# importar/visualizar a base de dados
tabela_busca = pd.read_excel('buscas.xlsx')
print(tabela_busca)

email = 'seu_email_aqui_@email.com'
tabela_ofertas = buscar_ofertas_produtos(driver, tabela_busca)
print(tabela_ofertas)
enviar_ofertas_email(tabela_ofertas, email)
