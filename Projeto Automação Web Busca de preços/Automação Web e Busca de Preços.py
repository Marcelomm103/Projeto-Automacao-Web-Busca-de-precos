#!/usr/bin/env python
# coding: utf-8

# In[17]:


# importar bibliotecas
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

import pandas as pd

import time

# criar o navegador
nav = webdriver.Chrome()

# importar/tratar/visualizar a base de dados
tabela_produtos = pd.read_excel("buscas.xlsx")
display(tabela_produtos)

tabela_produtos.loc[0, 'Preço mínimo'] = 4000
tabela_produtos.loc[1, 'Preço mínimo'] = 5000
tabela_produtos.loc[0, 'Preço máximo'] = 5000
tabela_produtos.loc[1, 'Preço máximo'] = 6000
display(tabela_produtos)


# Definição das funções de busca no google e no buscape

# In[18]:


def busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    # tratar os valores da função
    preco_maximo = float(preco_maximo)
    preco_minimo = float(preco_minimo)
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_produto = produto.split(" ")
    
    
    # entrar no buscape
    nav.get("https://www.buscape.com.br/")
    
    # pesquisar pelo produto no buscape
    nav.find_element(By.CLASS_NAME, 'search-bar__text-box').send_keys(produto, Keys.ENTER)
    
    # pegar a lista de resultados da busca do buscape
    time.sleep(5)
    lista_resultados = nav.find_elements(By.CLASS_NAME, 'Cell_Content__1630r')
    
    # para cada resultado
    lista_ofertas = []
    for resultado in lista_resultados:
        try:
            preco = resultado.find_element(By.CLASS_NAME, 'CellPrice_MainValue__3s0iP').text
            nome = resultado.get_attribute('title')
            nome = nome.lower()
            link = resultado.get_attribute('href')
            
            # verificacao do nome - se no nome tem algum termo banido
            tem_termos_banidos = False
            for palavra in lista_termos_banidos:
                if palavra in nome:
                    tem_termos_banidos = True  
                    
            # verificar se no nome tem todos os termos do nome do produto
            tem_todos_termos_produto = True
            for palavra in lista_termos_produto:
                if palavra not in nome:
                    tem_todos_termos_produto = False            
            
            if not tem_termos_banidos and tem_todos_termos_produto:
                preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
                preco = float(preco)
                if preco_minimo <= preco <= preco_maximo:
                    lista_ofertas.append((nome, preco, link))
        except:
            pass
    return lista_ofertas


# Contrução da nossa lista de ofertas encontradas

# In[19]:


tabela_ofertas = pd.DataFrame()

for linha in tabela_produtos.index:
    produto = tabela_produtos.loc[linha, "Nome"]
    termos_banidos = tabela_produtos.loc[linha, "Termos banidos"]
    preco_minimo = tabela_produtos.loc[linha, "Preço mínimo"]
    preco_maximo = tabela_produtos.loc[linha, "Preço máximo"]
        
    lista_ofertas_buscape = busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns=['produto', 'preco', 'link'])
        tabela_ofertas = tabela_ofertas.append(tabela_buscape)
    else:
        tabela_buscape = None

display(tabela_ofertas)


# Exportar a base de ofertas para Excel

# In[20]:


# exportar pro excel
tabela_ofertas = tabela_ofertas.reset_index(drop=True)
tabela_ofertas.to_excel("Ofertas.xlsx", index=False)


# Enviando o e-mail

# In[21]:


# enviar por e-mail o resultado da tabela
import win32com.client as win32

#verificando se existe alguma oferta dentro da tabela de ofertas
if len(tabela_ofertas.index) > 0:
    # vou enviar email
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'pythonimpressionador@gmail.com'
    mail.Subject = 'Produto(s) Encontrado(s) na faixa de preço desejada'
    mail.HTMLBody = f"""
    <p>Prezados,</p>
    <p>Encontramos alguns produtos em oferta dentro da faixa de preço desejada. Segue tabela com detalhes</p>
    {tabela_ofertas.to_html(index=False)}
    <p>Qualquer dúvida estou à disposição</p>
    <p>Att.,</p>
    """
    
    mail.Send()

nav.quit()  


# In[ ]:




