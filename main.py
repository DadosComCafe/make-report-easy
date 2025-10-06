from openpyxl import load_workbook
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


def get_numeric_columns(path: str) -> list:
    numeric_columns = []
    worksheet = load_workbook(path)
    worksheet = worksheet.active

    ipdb.set_trace()
    for col_idx in range(1, worksheet.max_column + 1):
        is_numeric_column = True
        for row_idx in range(1, worksheet.max_row + 1):
            cell_value = worksheet.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            if not isinstance(cell_value, (int, float)):
                is_numeric_column = False
                break

        if is_numeric_column:
            numeric_columns.append(col_idx)

if __name__ == "__main__":
    path = "assets/sample1.xlsx"
    logging.info("Encontrando colunas numéricas...")
    list_numeric_columns = get_numeric_columns(path=path)
    logging.info(f"Colunas numéricas: {list_numeric_columns}")

    #logging.info("Iniciando geração de relatório...")

    #create_only_numeric_sheet(path=path)

    #logging.info(f"Relatório pode ser acessado em {path}.")



