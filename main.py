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

