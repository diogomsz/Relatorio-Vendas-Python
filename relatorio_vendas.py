#!/usr/bin/env python
# coding: utf-8

# # Relatório de Vendas usando pandas e envio automático de e-mail com Python
# >Meu github: https://github.com/diogomsz
# >
# >ATENÇÃO: Aplicação feita 100% na IDE JupyterLab

# In[50]:


import pandas as pd
import win32com.client as win32

# importar a base de dados
tabela_vendas  = pd.read_excel("Vendas.xlsx")


# In[51]:


# visualizar a base de dados
display(tabela_vendas)


# In[52]:


# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
display(faturamento)


# In[53]:


# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
display(quantidade)


# In[54]:


# ticket medio por produto  em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
display(ticket_medio)


# In[55]:


# enviar um email com o relatorio
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'diogoassis3301@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,<p>

<p>Segue o Relatório de Vendas por cada Loja.<p>

<p>Faturamento:<p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:<p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:<p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2F}'.format})}

<p>Qualquer dúvida estou à disposição.<p>

<p>Att.,<p>
<p>Diogo<p>
'''

mail.Send()

