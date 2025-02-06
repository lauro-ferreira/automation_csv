import pandas as pd
import openpyxl


def format_csv_data(file, separador=","):
    """Formata os dados csv"""
    colunas_principais = ["Period", "Data_value", "STATUS", "Series_title_1", "UNITS"]
    df = pd.read_csv(file, sep=separador)

    # Apenas colunas principais
    df = df[colunas_principais]

    # Formato limpo de dados
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

    # Formatando "data_value"
    df["data_value"] = df["data_value"].apply(lambda x: f"${x:,.2f}")
    return df


def convert_csv_to_xlsx(file_csv, file_xlsx, separador=","):
    """
    Converte um CSV formatado para XLSX.

    Parâmetros:
    file_csv (str): Caminho do arquivo CSV de entrada.
    file_xlsx (str): Caminho do arquivo XLSX de saída.
    separador (str): Separador do CSV (padrão: ",").
    """
    # Formatar o CSV antes de converter
    df_formatado = format_csv_data(file_csv, separador)

    # Salvar como Excel (XLSX)
    df_formatado.to_excel(file_xlsx, index=False)
    print(f"Arquivo salvo como: {file_xlsx}")


convert_csv_to_xlsx("statistics-central-government.csv", "dados_formatados.xlsx")
