from flask import Flask, render_template, request, redirect, url_for
import os
import traceback
import time
import requests
import warnings
import urllib3
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from bs4 import BeautifulSoup

app = Flask(__name__)

warnings.filterwarnings("ignore", category=urllib3.exceptions.InsecureRequestWarning)

def consulta_mercado_livre(produto, maximotentativas=5):
    tentativas = 0
    if not " " in produto:
        produto = produto + " "

    produto = str(produto).replace(" ", "+")

    while tentativas <= maximotentativas:
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

            chrome_options = Options()
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--disable-gpu")
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)
            driver.get(location)

            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, 'ui-search-layout__item'))
            )

            action = ActionChains(driver)
            product_count = len(driver.find_elements(By.CSS_SELECTOR, '.poly-card__portada'))

            for i in range(product_count):
                try:
                    item = driver.find_elements(By.CSS_SELECTOR, '.poly-card__portada')[i]
                    action.move_to_element(item).perform()
                    time.sleep(0.5)
                except Exception as e:
                    print(f"[AVISO] Erro ao mover para item {i}: {e}")
                    continue

            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, 'ui-search-layout__item'))
            )

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
                if isinstance(preco, str):
                    preco = preco.replace('R$', '').replace('.', '').replace(',', '.').strip()
                    product['Preço'] = float(preco) if preco else 0.0

            produtos_ordenados = sorted(produtos, key=lambda x: x['Preço'])
            return produtos_ordenados

        except Exception as e:
            print("ERRO DE CONSULTA DO MERCADO LIVRE!!!!")
            print(f"Erro: {e}")
            print(f"Tipo de erro: {type(e)}")
            print(f"Traceback completo: {traceback.format_exc()}")
            print("TENTANDO NOVAMENTE.")
            tentativas += 1
            time.sleep(5)
            continue

    if tentativas >= maximotentativas:
        print("NÚMERO MÁXIMO DE TENTATIVAS")
    return None

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        produto = request.form['produto']
        return redirect(url_for('results', produto=produto))
    return render_template('index.html')

@app.route('/results')
def results():
    produto = request.args.get('produto')
    produtos = consulta_mercado_livre(produto)
    return render_template('lista-itens.html', produtos=produtos)

if __name__ == '__main__':
    app.run(debug=True)