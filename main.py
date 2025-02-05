# reads a spreadsheet
# formats the data
# and generates a report automatically
import pandas as pd
import openpyxl


def csv_para_excel(arquivo_csv, arquivo_excel):
    """Converte um arquivo CSV para Excel"""
    table = pd.read_csv("statistics-central-government.csv")  # Carrega o arquivo csv
    table.to_excel(arquivo_excel, index=False)  # Salva como excel

    print(f"Convertido com sucesso:{arquivo_excel}")


csv_para_excel("statistics.csv", "relatorio.xlsx")
