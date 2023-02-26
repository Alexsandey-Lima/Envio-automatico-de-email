import win32com.client as win32
import pandas as pd
import openpyxl



#criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

# criar um email
email = outlook.CreateItem(0)
list_email = pd.read_excel('clientes.xlsx')

# configurar as informações do seu e-mail
email.To = list_email.iloc[0,1]
print(email.To)
email.Subject = "E-mail automático do Python"
email.HTMLBody = f"""
<p>esse teste está endereçado para {list_email.iloc[0,0]}</p>
"""


email.Send()
print(f"Email Enviado para {list_email.iloc[0,0]}")




#list_email.iloc[linha,coluna]
#linha 0 = nome; linha 1= e-mail