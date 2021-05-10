#!/usr/bin/env python
# coding: utf-8

# In[6]:


import pyautogui
import time
import pyperclip
#importamos as bibliotecas para uso deste programa 


# In[7]:


pyautogui.alert("Não use seu computador enquanto é usado a automação.")
pyautogui.PAUSE = 1
#pausa de segurança até o computador obter resposta
pyautogui.hotkey("ctrl","t")
#hotkey usa para escrever pelo teclado
link = "https://drive.google.com/drive/folders/1wRTFw0sUVBjRr4hW5U9LF7DjLixRyxym"
#fez a variavel link para evitar erros na transcrição
pyperclip.copy(link)
#para copiar tem que usar pyperclip
pyautogui.hotkey("ctrl","v")
#hotkey usa para escrever no teclado aqui usamos o comando de colar
pyautogui.press("enter")
#para uso apenas de uma tecla do teclado usa o comando press
time.sleep(5)
#pausa para o carregamento da página 
pyautogui.click(1112,338)
pyautogui.click(1157,167)
pyautogui.click(991,554)
#pyautogui.click é usado para o mouse(descobrir posição na tela é o comando pyautogui.position()) caso queira click duplo acrescente click = 2
time.sleep(5)

import pandas as pd
#importamos a biblioteca pandas e colocamos apelido nela "pd" usando o comanda "as"
tabela = pd.read_excel(r"C:\Users\guilh\Downloads\Vendas - Dez.xlsx")
# criamos a variavel tabela, pedimos para o pandas ler e insermimos o endereço em que o arquivo está armazenado. Usamos o "r" antes do endereço para não ter alteração na transcrisição
display(tabela)
#comando para demosntrar a tabela
faturamento = tabela["Valor Final"].sum()
qtd_produtos = tabela["Quantidade"].sum()
#criamos duas variaveis para realizar a soma de suas colunas, o comando "sum()" realiza a soma dos produtos da coluna identificada no colchete

pyautogui.PAUSE = 1
#pausa de segurança até o computador obter resposta
pyautogui.hotkey("ctrl","t")
#hotkey usa para escrever pelo teclado
link = "https://outlook.live.com/owa/"
#fez a variavel link para evitar erros na transcrição
pyperclip.copy(link)
#para copiar tem que usar pyperclip
pyautogui.hotkey("ctrl","v")
#hotkey usa para escrever no teclado aqui usamos o comando de colar
pyautogui.press("enter")
#para uso apenas de uma tecla do teclado usa o comando press
time.sleep(3)
pyautogui.click(1084,95)
time.sleep(2)
pyautogui.click(808, 424)
pyautogui.click(799, 493)
time.sleep(4)
pyautogui.click(119,144)
#aqui foram os comandos para abrir o email
time.sleep(3)
pyautogui.write("guilherme+diretoria@live.it")
pyautogui.press("tab")
time.sleep(2)
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.press("tab")
time.sleep(2)
assunto = "Relatórios de Vendas"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
texto_email= f'''
Bom dia

O faturamento até ontem foi de R${faturamento:,.2f}
O estoque possui {qtd_produtos:,}

Att.:Guilherme Andrade

'''
pyperclip.copy(texto_email)
pyautogui.hotkey("ctrl","V")
time.sleep(2)
pyautogui.click(295,593)

pyautogui.alert("Finalizada a função. Liberado para uso!")


# In[ ]:





# 

# In[ ]:





# 

# In[ ]:





# 

# 

# In[ ]:





# In[4]:


time.sleep(5)
pyautogui.position()


# In[ ]:




