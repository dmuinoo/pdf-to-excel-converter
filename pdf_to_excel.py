import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

def extract_tables_from_pdf(pdf_path):
    """
    Extrae las tablas de un archivo PDF utilizando pdfplumber.

    Args:
        pdf_path (str): Ruta del archivo PDF.

    Returns:
        list: Lista de tablas extraídas como listas de listas.
    """
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables.extend(page.extract_tables())
    return tables

def combine_tables(tables):
    """
    Combina múltiples tablas en un solo DataFrame.

    Args:
        tables (list): Lista de DataFrames que representan las tablas.

    Returns:
        pandas.DataFrame: DataFrame combinado.
    """
    combined_table = pd.concat([pd.DataFrame(table) for table in tables], ignore_index=True)
    return combined_table

def split_columns_by_newline(table):
    """
    Divide las columnas de un DataFrame en nuevas columnas basadas en el separador '\n' en cada celda.

    Args:
        table (pandas.DataFrame): DataFrame que contiene las columnas a dividir.

    Returns:
        pandas.DataFrame: DataFrame con las columnas divididas.
    """
    new_columns = {}
    for column in table.columns:
        new_column = table[column].str.split("\n", expand=True)
        num_new_columns = new_column.shape[1]
        new_column_names = [f"{column}_{i}" for i in range(num_new_columns)]
        new_column.columns = new_column_names
        new_columns[column] = new_column
    combined_table = pd.concat(new_columns.values(), axis=1)
    return combined_table

def remove_rows_starting_with_puesto(table):
    """
    Elimina todas las filas que comienzan con "PUESTO" excepto la primera.

    Args:
        table (pandas.DataFrame): DataFrame que contiene las filas a filtrar.

    Returns:
        pandas.DataFrame: DataFrame con las filas filtradas.
    """
    first_puesto_index = table.index[table.iloc[:, 0].str.startswith("PUESTO")].min()
    table = pd.concat([table.iloc[[first_puesto_index]], table.iloc[first_puesto_index+1:].loc[~table.iloc[first_puesto_index+1:, 0].str.startswith("PUESTO")]], ignore_index=True)
    return table

def apply_bold_font_to_first_row(excel_file):
    """
    Aplica el formato en negrita a la primera fila de un archivo Excel.

    Args:
        excel_file (str): Ruta del archivo Excel.

    Returns:
        None
    """
    workbook = load_workbook(excel_file)
    sheet = workbook.active
    for cell in sheet["1"]:
        cell.font = Font(bold=True)
    workbook.save(excel_file)

# Ruta del archivo PDF que se va a procesar
pdf_path = "listado_puestos_GSIL.pdf"

# Extraer tablas del PDF
tables = extract_tables_from_pdf(pdf_path)

# Combinar tablas en un solo DataFrame
combined_table = combine_tables(tables)

# Dividir columnas basadas en '\n'
combined_table = split_columns_by_newline(combined_table)

# Eliminar filas que comienzan con "PUESTO" excepto la primera
combined_table = remove_rows_starting_with_puesto(combined_table)

# Ruta de salida del archivo Excel
output_excel_path = "output.xlsx"

# Guardar el DataFrame excluyendo la primera fila como un archivo Excel
combined_table.to_excel(output_excel_path, index=False, header=False)

# Aplicar negrita a la primera fila del archivo Excel
apply_bold_font_to_first_row(output_excel_path)
