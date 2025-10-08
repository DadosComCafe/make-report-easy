import logging
from random import randint
from typing import List

import ipdb
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

logging.basicConfig(level=logging.DEBUG)


def is_list_string(input_list: List) -> bool:
    return all(isinstance(item, str) for item in input_list)


def is_list_numeric(input_list: list) -> bool:
    """Checa se todos os elementos da lista `input_list` são numéricos (integers ou floats).

    Args:
        input_list (list): A lista que será checada.

    Returns:
        bool: True se todos os elementos forem numéricos, False se não forem.
    """

    return all(isinstance(item, (int, float)) for item in input_list)


def get_numeric_columns(path: str) -> List[int]:
    """Encontra todas as colunas que representam apenas valores numéricos.
    Para definir se a coluna é numérica, usa-se o primeiro, o último, e um valor de posição
    aleatória dentre estes intervalos, a fim de garantir que a coluna seja de fato numérica sem
    necessariamente processar a coluna inteira.

    Args:
        path (str): O caminho do arquivo xlsx que será examinado

    Returns:
        List[int]: Uma lista com as posições de colunas numéricas do arquivo xlsx analisado.
    """

    dict_of_lists = {}
    list_of_ids = []
    worksheet = load_workbook(path)
    worksheet = worksheet.active
    col_idx = 1

    logging.info("Coletando valores das colunas...")
    for i in range(col_idx, worksheet.max_column + 1):
        min_index = worksheet.min_row + 1
        max_index = worksheet.max_row
        random = randint(min_index + 2, max_index + 1) - 1
        column_values = [
            worksheet.cell(row=row_idx, column=i).value
            for row_idx in [
                worksheet.min_row + 1,
                worksheet.min_row + random,
                worksheet.max_row,
            ]
        ]
        dict_of_lists.update({i: column_values})

    logging.info("Verificando o tipo dos valores...")
    logging.info(f"Dicio: {dict_of_lists}")

    for key, value in dict_of_lists.items():
        if is_list_numeric(value):
            list_of_ids.append(key)
    return list_of_ids


def get_string_columns(path: str) -> List[int]:
    dict_of_lists = {}
    list_of_ids = []
    worksheet = load_workbook(path)
    worksheet = worksheet.active
    col_idx = 1

    logging.info("Coletando valores das colunas...")
    for i in range(col_idx, worksheet.max_column + 1):
        min_index = worksheet.min_row + 1
        max_index = worksheet.max_row
        random = randint(min_index + 2, max_index + 1) - 1
        column_values = [
            worksheet.cell(row=row_idx, column=i).value
            for row_idx in [
                worksheet.min_row + 1,
                worksheet.min_row + random,
                worksheet.max_row,
            ]
        ]
        dict_of_lists.update({i: column_values})

        logging.info("Verificando o tipo dos valores...")
        logging.info(f"Dicio: {dict_of_lists}")

        for key, value in dict_of_lists.items():
            if is_list_string(value):
                list_of_ids.append(key)
    return set(list_of_ids)


def create_only_numeric_sheet(path: str) -> None:
    """Exporta um novo xlsx com todo o conteúdo do arquivo informado no path,
    adicionado um sheet (planilha) contendo apenas as colunas numéricas do
    arquivo enviado.

    Args:
        path (str): O caminho do arquivo xlsx com o nome a extensão. Exemplo:
            assets/sample.xlsx
    """
    # ipdb.set_trace()
    # TODO: Terminar essa função, utilizando as funções anteriores criadas!
    numeric_new_file = path.replace(".xlsx", "_numeric.xlsx")

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
            if all(
                isinstance(cell.value, (int, float)) or cell.row == 1 for cell in col
            ):
                for cell in col:
                    numeric_sheet.cell(
                        row=cell.row, column=cell.column, value=cell.value
                    )

    wb_new.save(numeric_new_file)
    wb_new.save(numeric_new_file)


if __name__ == "__main__":
    path = "assets/file_sample.xlsx"
    logging.info("Encontrando colunas numéricas...")
    list_numeric_columns = get_numeric_columns(path=path)
    logging.info(f"Colunas numéricas: {list_numeric_columns}")

    logging.info({"..." * 20})

    logging.info("Encontrando colunas string...")
    list_string_columns = get_string_columns(path=path)
    logging.info(f"Colunas de texto: {list_string_columns}")

    # logging.info("Iniciando geração de relatório...")

    # create_only_numeric_sheet(path=path)

    # logging.info(f"Relatório pode ser acessado em {path}.")
