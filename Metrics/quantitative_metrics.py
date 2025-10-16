from openpyxl import Workbook, load_workbook


def get_sheet(path: str) -> Workbook:
    workbook = load_workbook(path)
    return workbook


def generate_new_sheet(path: str):
    workbook = get_sheet(path=path)
    workbook.create_sheet("NumericMetrics")
    workbook.save(f"{path.replace('.xlsx', '_numeric_repor.xlsx')}")
