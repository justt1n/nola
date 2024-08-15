import os
from app.models.ProcessModel import FilterProcessModel
from app.models.ProductSheetModel import ProductSheetModel


def build_filer_process_from_file(path: str):
    if not os.path.exists(path):
        return
    txt = ""
    with open(path, 'r') as file:
        txt = file.read()
    lines = txt.strip().split('\n')

    # First two lines are for payment spreadsheet and sheet name
    payment_spreadsheet = lines[0].split(":", 1)[1].strip()
    payment_sheet_name = lines[1].split(":", 1)[1].strip()

    product_spreadsheets = {}

    # Remaining lines should come in groups of three
    index = 1
    for i in range(2, len(lines), 3):
        product_spreadsheet = lines[i].split(":", 1)[1].strip()
        product_sheet_name = lines[i + 1].split(":", 1)[1].strip()
        headers = [header.strip() for header in lines[i + 2].split(":", 1)[1].split(",")]

        # Use the index as the key
        product_spreadsheets[str(index)] = ProductSheetModel(product_spreadsheet, product_sheet_name, headers)
        index += 1

    return FilterProcessModel(product_spreadsheets, payment_spreadsheet, payment_sheet_name)


class FilterProcessService:
    def __init__(self, filter_process_model: FilterProcessModel):
        self.filter_process_model = filter_process_model

    def filter(self):
        path = "done/test.txt"
        model = build_filer_process_from_file(path)
        print(model)


