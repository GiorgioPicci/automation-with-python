#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# In[55]:


#!pip install pyautogui
#!pip install pyperclip


# In[56]:


import pyautogui
import pyperclip
import time 

pyautogui.PAUSE = 1
#Passo 1: Entra no sistema da empresa/ abrir o navegador  
pyautogui.press("winleft")
pyautogui.write("navegador opera")
pyautogui.press("enter")
#pyautogui.hotkey("ctrl","t")=#usado quando já está em uma página web. 

#Passo 2: Navegar no sistema da empresa e encontar a base de dados()
link= "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing"
pyperclip.copy(link)
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(5)

#Passo 3: Exportar a base de dados (Fazer o download)
pyautogui.click(452, 335, clicks=2)
time.sleep(1)
pyautogui.click(452, 335) #clicar no aqruivo
pyautogui.click(452, 335, button ='right')#clicar botão direito 
time.sleep(1)
pyautogui.click(656, 709)#botão download
time.sleep(5)#esperar o download acontecer


# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[57]:


#Passo 4: Importar a Base de dados para o Python 
import pandas as pd
tabela = pd.read_excel(r"G:\Vendas - Dez.xlsx")
display(tabela)


# In[58]:


#Passo 5: calcular os indicadores 
#faturamento: soma da coluna valor final
faturamento = tabela["Valor Final"].sum()

#quantidade de protudos : soma da coluna quantidade
Quantidade = tabela["Quantidade"].sum()

display(faturamento)
display(Quantidade)


# ### Vamos agora enviar um e-mail pelo gmail

# In[59]:


#Passo 6: Enviar um email para a diretoria com o relatório 
#abri aba email
pyautogui.hotkey("ctrl","t")
pyautogui.write("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")
time.sleep(5)
# entrar no escrever email
pyautogui.click(74, 198)
time.sleep(3)
#escrever o destinatario do email
pyautogui.write("giorgiokuhnn@gmail.com")
time.sleep(1)
pyautogui.press("tab")#selecionando o destinatario 
pyautogui.press("tab")#seleciona o campo assunto
#escrever assunto do email
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl","v")
pyautogui.press("tab")#passar paro o corpo de email
#escrever o email

texto= f""" 
prezado, bom dia 
O faturamento foi de : R$ {faturamento:,.2f}
a quantidade foi de : {Quantidade:,}
giorgio""" 
pyperclip.copy(texto)
pyautogui.hotkey("ctrl","v")
#enviar email
pyautogui.hotkey("ctrl","enter")


# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# In[60]:


time.sleep(7)
pyautogui.position()


# In[ ]:




