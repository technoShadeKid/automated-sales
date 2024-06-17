import os
from datetime import datetime
import pandas as pd
import win32com.client as win32

paths = "bases"
files = os.listdir(paths)
print(files)

your_name = "(Insira seu nome, será mostrado ao final do e-mail)"
table = pd.DataFrame()

for file_name in files:
    sales_table = pd.read_csv(os.path.join(paths, file_name))
    sales_table["Data de Venda"] = pd.to_datetime("01/01/1900") + pd.to_timedelta(sales_table["Data de Venda"],
                                                                                    unit="d")
    table = pd.concat([table, sales_table])

table = table.sort_values(by="Data de Venda")
table = table.reset_index(drop=True)
table.to_excel("Vendas.xlsx", index=False)

outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
email.To = "(e-mail para onde o relatório será enviado)"
today_date = datetime.today().strftime("%d/%m/%Y")
email.Subject = f"Relatório de Vendas {today_date}"
email.Body = f"""
Prezados,

Segue em anexo o Relatório de Vendas de {today_date} atualizado.
Qualquer coisa estou à disposição.
Abs,
{your_name}
"""

paths = os.getcwd()
anexo = os.path.join(paths, "Vendas.xlsx")
email.Attachments.Add(anexo)

email.Send()