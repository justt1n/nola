from typing import List, Dict


class ProductSheetModel:
    def __init__(self, product_spreadsheets: str, product_sheet_name: str, headers: List[str]):
        self.product_spreadsheets = product_spreadsheets
        self.product_sheet_name = product_sheet_name
        self.headers = headers

    def __repr__(self):
        return f"Spreadsheet(product_spreadsheets='{self.product_spreadsheets}', product_sheet_name='{self.product_sheet_name}', headers={self.headers})"
