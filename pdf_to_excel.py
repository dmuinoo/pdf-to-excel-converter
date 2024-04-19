import pdfplumber
import pandas as pd

def extract_tables_from_pdf(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables.extend(page.extract_tables())
    return tables

def combine_tables(tables):
    combined_table = pd.concat([pd.DataFrame(table) for table in tables], ignore_index=True)
    return combined_table

def split_columns_by_newline(table):
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
    first_puesto_index = table.index[table.iloc[:, 0].str.startswith("PUESTO")].min()
    table = pd.concat([table.iloc[[first_puesto_index]], table.iloc[first_puesto_index+1:].loc[~table.iloc[first_puesto_index+1:, 0].str.startswith("PUESTO")]], ignore_index=True)
    return table



pdf_path = "listado_puestos_GSIL.pdf"
tables = extract_tables_from_pdf(pdf_path)
combined_table = combine_tables(tables)
combined_table = split_columns_by_newline(combined_table)
combined_table = remove_rows_starting_with_puesto(combined_table)
combined_table.to_excel("output.xlsx", index=False)
