import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment


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


def generate_report(file_xlsx, report_xlsx):
    """
    Gera um relatório formatado a partir do arquivo XLSX.

    Parâmetros:
    file_xlsx (str): Caminho do arquivo XLSX formatado.
    report_xlsx (str): Caminho do arquivo de relatório XLSX.
    """

    # Carregar os dados formatados
    df = pd.read_excel(file_xlsx)

    # Criar um escritor Excel para aplicar formatação
    with pd.ExcelWriter(report_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Relatório", index=False)

        # Carregar a planilha para aplicar estilos
        workbook = writer.book
        sheet = workbook["Relatório"]

        # Ajustar cabeçalhos: negrito e centralizado
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

        # Ajustar largura das colunas automaticamente
        for column in sheet.columns:
            max_length = 0
            col_letter = column[0].column_letter
            for cell in column:
                try:
                    max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = max_length + 2
            sheet.column_dimensions[col_letter].width = adjusted_width

        # Salvar relatório formatado
        workbook.save(report_xlsx)
