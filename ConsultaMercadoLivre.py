import os
import traceback
from logging import exception
import time
import requests
import warnings
import urllib3
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from openpyxl.worksheet.hyperlink import Hyperlink
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pandas as pd
from bs4 import BeautifulSoup

warnings.filterwarnings("ignore", category=urllib3.exceptions.InsecureRequestWarning)

def ConsultaMercadoLivre(maximotentativas=5):
    tentativas = 0
    produto = ""

    while tentativas <= maximotentativas:

        while produto == "":
            produto = input("Digite um item para pesquisar no Mercado Livre: ")
            if not " " in produto:
               produto = produto + " "

            produto = str(produto).replace(" ", "+")

        try:
            url = 'https://www.mercadolivre.com.br/'
            head = {
              'Host': 'www.mercadolivre.com.br',
              'Connection': 'keep-alive',
              'sec-ch-ua': '"Not:A-Brand";v="24", "Chromium";v="134"',
              'sec-ch-ua-mobile': '?0',
              'sec-ch-ua-platform': '"Windows"',
              'Accept-Language': 'pt-BR,pt;q=0.9',
              'Upgrade-Insecure-Requests': '1',
              'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36',
              'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
              'Referer': 'https://www.google.com/',
              'Accept-Encoding': 'gzip, deflate, br, zstd',
            }
            response = requests.get(url, headers=head, verify=False)

            sessionId = response.cookies['_mldataSessionId']
            d2id = response.cookies['_d2id']
            navigation = response.cookies['c_ui-navigation']

            url = f'https://lista.mercadolivre.com.br/{produto}'
            head = {
              'Host': 'lista.mercadolivre.com.br',
              'Connection': 'keep-alive',
              'Accept-Language': 'pt-BR,pt;q=0.9',
              'Upgrade-Insecure-Requests': '1',
              'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36',
              'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
              'Sec-Fetch-Site': 'same-site',
              'Sec-Fetch-Mode': 'navigate',
              'Sec-Fetch-User': '?1',
              'Sec-Fetch-Dest': 'document',
              'sec-ch-ua': '"Not:A-Brand";v="24", "Chromium";v="134"',
              'sec-ch-ua-mobile': '?0',
              'sec-ch-ua-platform': '"Windows"',
              'Referer': 'https://www.mercadolivre.com.br/',
              'Cookie': f'_mldataSessionId={sessionId}; _d2id={d2id}',
            }
            response = requests.get(url, headers=head, verify=False, allow_redirects=False)

            location = response.headers['location']

            url = location
            head = {
              'Host': 'lista.mercadolivre.com.br',
              'Connection': 'keep-alive',
              'Accept-Language': 'pt-BR,pt;q=0.9',
              'Upgrade-Insecure-Requests': '1',
              'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36',
              'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
              'Sec-Fetch-Site': 'same-site',
              'Sec-Fetch-Mode': 'navigate',
              'Sec-Fetch-User': '?1',
              'Sec-Fetch-Dest': 'document',
              'device-memory': '8',
              'dpr': '1',
              'viewport-width': '1920',
              'rtt': '50',
              'downlink': '1.65',
              'ect': '4g',
              'Referer': 'https://www.mercadolivre.com.br/',
              'Cookie': f'_mldataSessionId={sessionId}; _d2id={d2id}; _csrf=tP7OUS8o1QLoRjSKBaXY7_v2; _mldataSessionId={sessionId}; c_ui-navigation={navigation}',
            }
            response = requests.get(url, headers=head, verify=False)

            # Caminho para o seu ChromeDriver
            driver_path = 'chromedriver-win64/chromedriver.exe'

            # Configurações do Selenium para rodar em segundo plano (sem abrir o navegador)
            chrome_options = Options()
            chrome_options.add_argument("--headless")  # Isso vai rodar o navegador em segundo plano
            chrome_options.add_argument("--disable-gpu")

            # Configura o serviço do ChromeDriver
            service = Service(driver_path)

            # Inicializa o WebDriver
            driver = webdriver.Chrome(service=service, options=chrome_options)

            # URL do Mercado Livre
            # Acessa a página
            driver.get(location)

            # Espera até que os produtos estejam visíveis (ajuste para o seu caso)
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, 'ui-search-layout__item'))
            )

            # Encontra todos os itens de produto (ajuste conforme necessário)
            product_items = driver.find_elements(By.CSS_SELECTOR, '.poly-card__portada')

            # Cria uma instância do ActionChains para mover o mouse
            action = ActionChains(driver)

            # Passa o mouse sobre cada item para garantir que tudo seja carregado
            for item in product_items:
                action.move_to_element(item).perform()
                time.sleep(0.5)
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, 'ui-search-layout__item'))
            )

                # Pega o HTML da página renderizada
            html = driver.page_source

            soup = BeautifulSoup(html, "html.parser")

            items = soup.find_all('li', class_='ui-search-layout__item')

            produtos = []

            for item in items:
                nome = item.find('a', class_='poly-component__title')
                if nome:
                    nome = nome.text.strip()

                imagem = item.find('img', class_='poly-component__picture')
                if imagem:
                    imagem = imagem['src']

                preco = item.find('span', class_='andes-money-amount andes-money-amount--cents-superscript')
                if preco:
                    preco = preco.text.strip()

                link = item.find('a', class_='poly-component__title')
                if link:
                    link = link['href']

                produtos.append({"Nome": nome, "Preço": preco, "imagem": imagem, "link": link})

            for product in produtos:
                preco = product['Preço']

                # Verificar se o preço é uma string antes de aplicar 'replace'
                if isinstance(preco, str):
                    # Remove qualquer coisa que não seja número ou ponto (por exemplo, R$, vírgulas)
                    preco = preco.replace('R$', '').replace('.', '').replace(',', '.').strip()
                    product['Preço'] = float(preco) if preco else 0.0

                # Ordenar a lista por preço de forma decrescente
            produtos_ordenados = sorted(produtos, key=lambda x: x['Preço'])

            # Criar um DataFrame do Pandas para poder salvar no Excel
            df = pd.DataFrame(produtos_ordenados)

            excel_file = 'produtos_ordenados.xlsx'
            # Salvar em um arquivo Excel
            df.to_excel('produtos_ordenados.xlsx', index=False)

            # Carregar o arquivo Excel usando openpyxl para formatar as células
            wb = load_workbook(excel_file)
            ws = wb.active

            # Criar um estilo de formatação de moeda
            currency_style = NamedStyle(name="currency_style", number_format='R$ #,##0.00')

            # Aplicar a formatação de moeda na coluna de preços (supondo que o preço esteja na coluna 'Preço' que é a coluna 3)
            for row in range(2, len(df) + 2):  # Começando da linha 2, já que a linha 1 é de cabeçalho
                cell = ws.cell(row=row, column=2)  # A coluna 3 é a coluna do preço
                cell.style = currency_style

            # Salvar o arquivo Excel com a formatação aplicada
            wb.save(excel_file)

            # Mensagem de sucesso
            print("CRIADO ARQUIVO EXCEL COM SUCESSO")
            os.startfile(excel_file)

            break

        except Exception as e:
            print("ERRO DE CONSULTA DO MERCADO LIVRE!!!!")
            print(f"Erro: {e}")  # Mostra a mensagem do erro
            print(
                f"Tipo de erro: {type(e)}")  # Mostra o tipo do erro (por exemplo, requests.exceptions.RequestException)
            print(
                f"Traceback completo: {traceback.format_exc()}")  # Mostra o traceback completo para entender onde o erro aconteceu
            print("TENTANDO NOVAMENTE.")
            tentativas+=1
            time.sleep(5)
            continue

    if tentativas >= maximotentativas:
        print("NÚMERO MÁXIMO DE TENTATIVAS")

consulta = ConsultaMercadoLivre()






