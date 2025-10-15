from base_sheets import BaseSheet


class NumericSheet(BaseSheet):
    def __init__(self, path: str):
        super().__init__(path)
        self.new_worksheet.title = "NumericColumns"

        self.get_numeric_columns = self.get_uniquetype_columns(unique_type="numeric")

    def create_only_numeric_sheet(self):
        return self.create_unique_type_sheet(unique_type="numeric")


class StringSheet(BaseSheet):
    def __init__(self, path: str):
        super().__init__(path)
        self.new_worksheet.title = "StringColumns"

        self.get_text_columns = self.get_uniquetype_columns(unique_type="text")

    def create_only_text_sheet(self):
        return self.create_unique_type_sheet(unique_type="text")


if __name__ == "__main__":
    path = "assets/file_sample.xlsx"
    numeric_sheet_obj = NumericSheet(path=path)
    print(numeric_sheet_obj.get_numeric_columns)
    numeric_sheet_obj.create_only_numeric_sheet()

    string_sheet_obj = StringSheet(path=path)
    print(string_sheet_obj.get_text_columns)
    string_sheet_obj.create_only_text_sheet()
