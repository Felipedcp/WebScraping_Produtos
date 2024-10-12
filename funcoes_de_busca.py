from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import win32com.client as win32
from datetime import datetime
import pandas as pd
import pythoncom
import pytz
import time

sao_paulo_tz = pytz.timezone('America/Sao_Paulo')

def verificar_tem_termos_banidos(lista_termos_banidos: list[str], nome_oferta: str) -> bool:
        """Verifica se tem qualquer termo banido no nome da oferta encontrada

        Parameters:
            lista_termos_banidos (list[str]): lista que contém todos os termos banidos
            nome_oferta (str): nome da oferta encontrada
        
        Returns:
            tem_termos_banidos (bool): True se tem qualquer termo banido no nome da oferta encontrada, False caso contrário
        
        """

        tem_termos_banidos = False

        # Se a lista_termos_banidos é igual a None, isto significa que o campo "termos banidos" não foi preenchido.
        if lista_termos_banidos == None :
            return False
        
        for palavra in lista_termos_banidos:
            if palavra in nome_oferta:
                tem_termos_banidos = True
        return tem_termos_banidos        


def verificar_tem_todos_termos_produto(lista_termos_nome_produto: list[str], nome_oferta: str) -> bool:
        """Verifica se todos os termos do produto estão presentes no nome da oferta encontrada
        
        Parameters:
            lista_termos_nome_produto (list[str]): lista que contém todos os termos do produto
            nome_oferta (str): nome da oferta encontrada

        Returns:
            tem_todos_termos_produto (bool): True se todos os termos do produto estão na oferta encontrada, False caso contrário
        """

        tem_todos_termos_produto = True
        for palavra in lista_termos_nome_produto:
            if palavra not in nome_oferta:
                tem_todos_termos_produto = False
        return tem_todos_termos_produto


def busca_google_shopping(driver, produto: str, termos_banidos: str, preco_minimo: float , preco_maximo: float) -> pd.DataFrame:
    """Faz uma busca no google shopping pelo produto desejado e filtra os resultados de acordo com os termos banidos, preço mínimo e preço máximo do produto
        
    Parameters:
        driver: webdriver que controla o browser
        produto (str): nome do produto
        termos_banidos (str): termos que não poderão aparecer no nome do produto encontrado
        preco_minimo (float): preço mínimo do produto
        preco_maximo (float): preço máximo do produto

    Returns:
        df_ofertas (pd.DataFrame): Dataframe com as colunas Produto, Preço, Tipo, Data, Hora e Link contendo todas as ofertas encontradas 
    """

    # tratando nome do produto, termos_banidos, preco_minimo e preco_maximo
    produto = produto.lower()
    lista_termos_nome_produto = produto.split(" ")

    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")

    if len(termos_banidos) == 0:
        termos_banidos = None
        lista_termos_banidos = None

    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)
    
    df_ofertas = pd.DataFrame({
        'Produto': [],
        'Preço': [],
        'Tipo': [],
        'Data': [],
        'Hora': [],
        'Link': [],
    })

    # entrar no google shopping
    driver.get('https://shopping.google.com/')
    driver.find_element(By.XPATH, '//*[@id="REsRA"]').send_keys(produto, Keys.ENTER)
    time.sleep(2)

    # pegar as informações do produto
    lista_resultados = driver.find_elements(By.CLASS_NAME, 'i0X6df')
    for resultado in lista_resultados:
        nome = resultado.find_element(By.CLASS_NAME, 'tAxDx').text
        nome = nome.lower()

        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)
        tem_todos_termos_produto = verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome)

        # selecionar só os elementos que tem_termos_banidos = False e tem_todos_termos_produto = True
        if not tem_termos_banidos and tem_todos_termos_produto:
            # tratar o preço e o converter para float
            preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
            preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            try:
                preco = float(preco)
            except ValueError:
                continue

            # verificar se o preço está entre o preco_minimo e o preco_maximo
            if preco_minimo <= preco <= preco_maximo:
                now = datetime.now(sao_paulo_tz)
                data = now.strftime('%d/%m/%Y')
                hora = now.strftime('%H:%M')
                link = resultado.find_element(By.CLASS_NAME, 'Lq5OHe').get_attribute('href')
                df_ofertas.loc[len(df_ofertas.index)] = [nome, preco, produto, data, hora, link]

    return df_ofertas


def busca_buscape(driver, produto: str, termos_banidos: str, preco_minimo: float, preco_maximo: float) -> pd.DataFrame:
    """Faz uma busca no buscapé pelo produto desejado e filtra os resultados de acordo com os termos banidos, preço mínimo e preço máximo do produto
        
    Parameters:
        driver: webdriver que controla o browser
        produto (str): nome do produto
        termos_banidos (str): termos que não poderão aparecer no nome do produto encontrado
        preco_minimo (float): preço mínimo do produto
        preco_maximo (float): preço máximo do produto

    Returns:
        df_ofertas (pd.DataFrame): Dataframe com as colunas Produto, Preço, Tipo, Data, Hora e Link contendo todas as ofertas encontradas 
    """

    # tratando nome do produto, termos_banidos, preco_minimo e preco_maximo
    produto = produto.lower()
    lista_termos_nome_produto = produto.split(" ")

    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")

    if len(termos_banidos) == 0:
        termos_banidos = None
        lista_termos_banidos = None

    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)
    
    df_ofertas = pd.DataFrame({
        'Produto': [],
        'Preço': [],
        'Tipo': [],
        'Data': [],
        'Hora': [],
        'Link': [],
    })

    # entrar no buscapé
    driver.get('https://www.buscape.com.br/')
    driver.find_element(By.XPATH, '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(produto, Keys.ENTER)
    time.sleep(2)

    lista_resultados = driver.find_elements(By.CLASS_NAME, 'ProductCard_ProductCard_Inner__gapsh')
    for resultado in lista_resultados:
        nome = resultado.find_element(By.TAG_NAME, 'h2').text
        nome = nome.lower()

        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)
        tem_todos_termos_produto = verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome)
        
        # selecionar só os elementos que tem_termos_banidos = False e tem_todos_termos_produto = True
        if not tem_termos_banidos and tem_todos_termos_produto:
            # tratar o preço e o converter para float
            preco = resultado.find_element(By.CLASS_NAME, 'Text_MobileHeadingS__HEz7L').text
            preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            try:
                preco = float(preco)
            except ValueError:
                continue

            # verificar se o preço está entre o preco_minimo e o preco_maximo
            if preco_minimo <= preco <= preco_maximo:
                now = datetime.now(sao_paulo_tz)
                data = now.strftime('%d/%m/%Y')
                hora = now.strftime('%H:%M')
                link = resultado.get_attribute("href")
                df_ofertas.loc[len(df_ofertas.index)] = [nome, preco, produto, data, hora, link]
    
    return df_ofertas


def enviar_ofertas_email(tabela_ofertas: pd.DataFrame, receiver_email: str):
    """Envia uma mensagem com a tabela de ofertas para o e-mail do destinatário através do outlook
        
    Parameters:
        tabela_ofertas (pd.DataFrame): tabela resultado da busca de produtos
        receiver_email (str): e-mail do destinatário

    """

    # Envia email com tabela de ofertas
    if len(tabela_ofertas.index) > 0:
             
        # Enviar por e-mail o resultado da tabela
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = receiver_email
        mail.Subject = 'Produto(s) Encontrado(s) na faixa de preço desejada'
        mail.HTMLBody = f"""
        <p>Prezados,</p>
        <p>Encontramos alguns produtos em oferta dentro da faixa de preço desejada</p>
        {tabela_ofertas.to_html(index=False)}
        <p>Att.,</p>
        """
        mail.Send()
        pythoncom.CoUninitialize()


def buscar_ofertas_produtos(driver, tabela_busca: pd.DataFrame) -> pd.DataFrame:
    """Faz busca no google shopping e no buscapé pelo produto desejado, e se o arquivo "tabela_ofertas.xlsx" já existir, adiciona o resultado da busca
    nesse arquivo, caso contrário, cria o arquivo "tabela_ofertas.xlsx" com o resultado da busca.
        
    Parameters:
        driver: webdriver que controla o browser
        tabela_busca (pd.DataFrame): Dataframe com as colunas Nome, Termos banidos, Preço mínimo e Preço máximo, contendo as informações de busca dos produtos

    Returns:
        tabela_ofertas (pd.DataFrame): Dataframe com as colunas Produto, Preço, Tipo, Data, Hora e Link contendo todas as ofertas de cada produto em ordem ascendente de preço
    """
    
    tabela_ofertas = pd.DataFrame(columns=['Produto', 'Preço', 'Tipo', 'Data', 'Hora', 'Link'])

    # pesquisar pelo produto
    for linha in tabela_busca.index:
        tabela_temporaria = pd.DataFrame()
        produto = tabela_busca.loc[linha, "Nome"]
        termos_banidos = tabela_busca.loc[linha, "Termos banidos"]
        preco_minimo = float(tabela_busca.loc[linha, "Preço mínimo"])
        preco_maximo = float(tabela_busca.loc[linha, "Preço máximo"])
        
        # busca pelo produto no google shopping
        df_ofertas_google_shopping = busca_google_shopping(driver, produto, termos_banidos, preco_minimo, preco_maximo)
        if len(df_ofertas_google_shopping.index) > 0:
            tabela_temporaria = pd.concat([tabela_temporaria, df_ofertas_google_shopping], ignore_index=True)

        # busca pelo produto no buscapé
        df_ofertas_buscape = busca_buscape(driver, produto, termos_banidos, preco_minimo, preco_maximo)
        if len(df_ofertas_buscape.index) > 0:
            tabela_temporaria = pd.concat([tabela_temporaria, df_ofertas_buscape], ignore_index=True)

        # Se tabela_temporaria não estiver vazia, ordena as ofertas em ordem crescente de preço e as adiciona na tabela_ofertas.
        if len(tabela_temporaria.index) > 0:
            tabela_temporaria = tabela_temporaria.sort_values(by=['Preço'], ignore_index=True)
            tabela_ofertas = pd.concat([tabela_ofertas, tabela_temporaria], ignore_index=True)
    
    # exportar tabela para arquivo do tipo .xlsx
    if len(tabela_ofertas.index) > 0:
        try:
            planilha_ofertas = pd.read_excel("tabela_ofertas.xlsx")
            planilha_ofertas = pd.concat([planilha_ofertas, tabela_ofertas], ignore_index=True)
            planilha_ofertas.to_excel("tabela_ofertas.xlsx", index=False)
        except:
            tabela_ofertas.to_excel("tabela_ofertas.xlsx", index=False)

    driver.close()
    return tabela_ofertas