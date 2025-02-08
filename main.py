import smtplib
import pandas as pd
from openpyxl.styles import Font, Alignment
from email.message import EmailMessage


def format_csv_data(file, separator=","):
    """Formats CSV data"""
    main_columns = ["Period", "Data_value", "STATUS", "Series_title_1", "UNITS"]
    df = pd.read_csv(file, sep=separator)

    # Only main columns
    df = df[main_columns]

    # Clean data format
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

    # Formatting "data_value"
    df["data_value"] = df["data_value"].apply(lambda x: f"${x:,.2f}")
    return df


def convert_csv_to_xlsx(csv_file, xlsx_file, separator=","):
    """Converts a formatted CSV to XLSX."""
    # Format the CSV before converting
    formatted_df = format_csv_data(csv_file, separator)

    # Save as Excel (XLSX)
    formatted_df.to_excel(xlsx_file, index=False)
    print(f"File saved as: {xlsx_file}")


convert_csv_to_xlsx("statistics-central-government.csv", "file.xlsx", ",")


def generate_report(xlsx_file, report_xlsx):
    """Generates a formatted report from the XLSX file."""

    # Load formatted data
    df = pd.read_excel(xlsx_file)

    # Create an Excel writer to apply formatting
    with pd.ExcelWriter(report_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Report", index=False)

        # Load the worksheet to apply styles
        workbook = writer.book
        sheet = workbook["Report"]

        # Adjust headers: bold and centered
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

        # Adjusting column width
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

        # Save formatted report
        workbook.save(report_xlsx)


generate_report("file.xlsx", "report.xlsx")


email = "your email"
with open("password.txt") as f:
    password = f.readlines()

    f.close()

    email_password = password[0]

msg = EmailMessage()
msg["Subject"] = "Central Government Statistics Report"
msg["From"] = "your email"
msg["To"] = "recipient email"
msg.set_content("Attached is the statistics report:")

with open("report.xlsx", "rb") as re:
    content = re.read()
    msg.add_attachment(content, maintype="application", subtype="xlsx", filename="report.xlsx")

with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login(email, email_password)
    smtp.send_message(msg)
