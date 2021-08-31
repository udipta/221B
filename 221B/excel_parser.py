import openpyxl
import os


class ExcelParser:
    """
        Functionality to process the excel data file and generate json data for specific tags
    """

    def __init__(self, input_excel_file: str = 'ToParse_Python.xlsx'):
        self.input_excel_file = input_excel_file
        self.input_excel_workbook = openpyxl.load_workbook(self.input_excel_file, data_only=True) if os.path.isfile(
            self.input_excel_file) else None

        self.sheet = None
        self.output_json = []

    def sheet_parser(self):
        order_dict = {
            'Quote Number': '',
            'Date': '',
            'Ship To': '',
            'Ship From': '',
            'Name': ''}

        item_dict = {
            'LineNumber': '',
            'PartNumber': '',
            'Description': '',
            'Item Type': '',
            'Price': ''}

        row_start, row_end, col_start, col_end = 0, 0, 0, 0
        for row in self.sheet.rows:
            for cell in row:
                if "Name" in str(cell.value):
                    k, v = cell.value.split(':')
                    order_dict[k] = v
                elif cell.value in order_dict.keys():
                    order_dict[cell.value] = row[cell.column].value.date().isoformat() if cell.value == "Date" else row[cell.column].value
                elif cell.value in item_dict.keys():
                    if cell.value == "LineNumber":
                        row_start = cell.row + 1
                        col_start = cell.column - 1
                    elif cell.value == "Price":
                        col_end = cell.column
                elif cell.value == "-----------":
                    row_end = cell.row - 1

        item_list = []
        for i in range(row_start, row_end + 1):
            row = [cell.value for cell in self.sheet[i][col_start:col_end + 1]]
            if any(row):
                item_list.append(dict(zip(item_dict.keys(), row)))

        order_dict.update({'Items': item_list})

        return order_dict

    def parser(self):
        for sheet in self.input_excel_workbook:
            self.sheet = sheet
            output = self.sheet_parser()
            if output:
                self.output_json.append(output)


if __name__ == '__main__':
    try:
        ep = ExcelParser()
        ep.parser()
        print(ep.output_json)
    except Exception as e:
        import traceback
        print(traceback.format_exc())
