from openpyxl.utils import column_index_from_string
from openpyxl import load_workbook
from typing import List
from random import randint
import logging
import ipdb

logging.basicConfig(level=logging.DEBUG)


def create_only_numeric_sheet(path: str) -> None:
    """Exporta um novo xlsx com todo o conteúdo do arquivo informado no path,
    adicionado um sheet (planilha) contendo apenas as colunas numéricas do 
    arquivo enviado.

    Args:
        path (str): O caminho do arquivo xlsx com o nome a extensão. Exemplo:
            assets/sample.xlsx
    """
    #ipdb.set_trace()
    numeric_new_file = path.replace('.xlsx', '_numeric.xlsx')

    wb_original = load_workbook(path)

    wb_original.save(numeric_new_file)
    wb_new = load_workbook(numeric_new_file)

    sheets = wb_original.sheetnames

    for sheet_name in sheets:
        numeric_sheet_name = f"numeric_{sheet_name}"
        if numeric_sheet_name not in wb_new.sheetnames:
            wb_new.create_sheet(numeric_sheet_name)
        
        old_sheet = wb_original[sheet_name]
        numeric_sheet = wb_new[numeric_sheet_name]
        
        for col in old_sheet.iter_cols():
            if all(isinstance(cell.value, (int, float)) or cell.row == 1 for cell in col):
                for cell in col:
                    numeric_sheet.cell(row=cell.row, column=cell.column, value=cell.value)

    wb_new.save(numeric_new_file)

    #for col in reversed(list(numeric_sheet.iter_cols(min_row=1, max_row=numeric_sheet.max_row, min_col=1, max_col=numeric_sheet.max_column))):
    #    if all(cell.value is None for cell in col):
    #        numeric_sheet.delete_cols(col[0].column)

    wb_new.save(numeric_new_file)

def is_list_numeric(input_list):
    """
    Checks if all elements in a list are numeric (integers or floats).

    Args:
        input_list: The list to check.

    Returns:
        True if all elements are numeric, False otherwise.
    """
    return all(isinstance(item, (int, float)) for item in input_list)

def get_numeric_columns(path: str) -> List[int]:
    dict_of_lists = {}
    list_of_ids = []
    worksheet = load_workbook(path)
    worksheet = worksheet.active
    col_idx = 1 #worksheet.min_column
    logging.info("Coletando valores das colunas...")
    for i in range(col_idx, worksheet.max_column+1):
        min_index = worksheet.min_row+1
        max_index = worksheet.max_row
        random = randint(min_index+2, max_index+1) - 1
        column_values = [worksheet.cell(row=row_idx, column=i).value for row_idx in [worksheet.min_row+1, worksheet.min_row+random,worksheet.max_row]]
        dict_of_lists.update({i: column_values})
    logging.info("Verificando o tipo dos valores...")
    logging.info(f"Dicio: {dict_of_lists}")
    for key, value in dict_of_lists.items():
        if is_list_numeric(value):
            list_of_ids.append(key)
    return list_of_ids



if __name__ == "__main__":
    path = "assets/file_sample.xlsx"
    logging.info("Encontrando colunas numéricas...")
    list_numeric_columns = get_numeric_columns(path=path)
    logging.info(f"Colunas numéricas: {list_numeric_columns}")

    #logging.info("Iniciando geração de relatório...")

    #create_only_numeric_sheet(path=path)

    #logging.info(f"Relatório pode ser acessado em {path}.")



