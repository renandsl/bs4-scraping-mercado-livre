from bs4 import BeautifulSoup
from selenium import webdriver
from time import sleep
from urllib.parse import urlparse
import openpyxl
import PySimpleGUI as sg
import pandas as pd

def coletar_dados(produtos, nomes, precos, links):
    
    for produto in produtos:
        nome = produto.find('h2', 'ui-search-item__title')

        if nome:
            nome = nome.text.strip()
        else:
            nome = 'N/A'

        preco_inteiro_elemento = produto.find('span', {'class': 'price-tag-fraction'})
        preco_centavos_elemento = produto.find('span', {'class': 'price-tag-cents'})

        if preco_inteiro_elemento and preco_centavos_elemento:
            preco_total = preco_inteiro_elemento.text.strip() + '.' + preco_centavos_elemento.text.strip()
        elif preco_inteiro_elemento:
            preco_total = preco_inteiro_elemento.text.strip() + '.00'
        else:
            preco_total = None

        link_elemento = produto.find('a', class_='ui-search-item__group__element')

        if link_elemento:
            link = link_elemento['href']
            parsed_link = urlparse(link)
            link = parsed_link.netloc + parsed_link.path 
        else:
            link = 'N/A'

        nomes.append(nome)
        precos.append(preco_total)
        links.append(link)

def formatar_preco(preco):

    try:
        preco_float = float(preco)
        return f'{preco_float:,.2f}'.replace('.', '-')
    except ValueError:
        return preco

def salvar_dados(nomes, precos, links):
    
    dados = {'Nome': nomes, 'Preço': precos, 'Link': links}
    df = pd.DataFrame(dados)

    df['Preço'] = df['Preço'].apply(formatar_preco)
    df['Preço'] = df['Preço'].str.replace(',', '.').str.replace('-', ',')

    df['Preço'] = df['Preço'].str.replace('.', '').str.replace(',', '.').astype(float)
    df = df.sort_values('Preço')

    df['Preço'] = df['Preço'].apply(lambda x: f'{x:.2f}')
    df.to_excel('base_de_dados.xlsx', index=False)

def redefinir_planilha(arquivo):

    workbook = openpyxl.load_workbook(arquivo)
    planilha = workbook.active
    planilha.delete_rows(1, planilha.max_row)
    workbook.save(arquivo)

def rodar_script(valores, num_paginas):

    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    driver = webdriver.Chrome(options=options)

    indice = 50
    referencia = int('01')
    pagina = 0

    url = f"https://lista.mercadolivre.com.br/{valores['pesquisa']}_Desde_{referencia}_NoIndex_True"

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    produtos = soup.find_all('div', 'ui-search-result__content-wrapper')

    nomes = []
    precos = []
    links = []
    
    for _ in range(1, num_paginas):

        url = f"https://lista.mercadolivre.com.br/{valores['pesquisa']}_Desde_{referencia}_NoIndex_True"
        referencia += indice
        pagina += 1
        driver.get(url)
        driver.maximize_window()
        sg.popup_auto_close(f'Fazendo scraping na página [{pagina}].', auto_close_duration=5)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        produtos = soup.find_all('div', 'ui-search-result__content-wrapper')
        coletar_dados(produtos, nomes, precos, links)
        sleep(1)
        salvar_dados(nomes, precos, links)

sg.theme('DarkTeal9')

layout = [
    [sg.Column([[sg.Image(r'C:\\Users\\cliente\\Py\\Scraping\\scrapingml.png')]], justification='center', pad=(0, 30))],
    [sg.Text('Nome do produto', justification='left', size=(20, 1)), sg.Input(key='pesquisa', size=(25, 1), justification='left')],
    [sg.Text('Número de páginas', justification='left', size=(20, 1)), sg.Input(key='num_paginas', size=(25, 1), justification='left')],
    [sg.Button('Iniciar Scraping', size=(20, 1), pad=(5, 10), tooltip='No máximo 40 páginas.', button_color=('white', '#0080FF'), border_width=2),
     sg.Button('Redefinir Planilha', size=(20, 1), pad=(5, 10), tooltip='Redefinir a planilha de dados', button_color=('white', '#0080FF'), border_width=2)],
]

janela = sg.Window('M.L Scraping', layout, size=(350, 310))

while True:
    evento, valores = janela.read()

    if evento == sg.WIN_CLOSED:
        janela.close()
        break

    if evento == 'Redefinir Planilha':
        redefinir_planilha('base_de_dados.xlsx')
        sg.popup('A planilha foi redefinida.')

    if evento == 'Iniciar Scraping':
        produto = valores['pesquisa']
        if produto.strip() == '':
            sg.popup('Você precisa digitar um produto.')
            continue

        num_paginas = valores['num_paginas']
        if num_paginas.strip() == '':
            sg.popup('Você precisa fornecer o número de páginas.')
            continue

        num_paginas = int(num_paginas)
        if num_paginas > 40:
            sg.popup('O valor ultrapassou as 40 páginas permitidas.')
            num_paginas = 40

        rodar_script(valores, num_paginas)
        sg.popup('Scraping finalizado')
        sleep(2)
        