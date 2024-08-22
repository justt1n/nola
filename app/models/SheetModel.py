from typing import List, Dict


class ProductSheetModel:
    def __init__(self, product_spreadsheets: str, product_sheet_name: str, headers: List[str], header_row_index: int):
        self.product_spreadsheets = product_spreadsheets
        self.product_sheet_name = product_sheet_name
        self.headers = headers
        self.header_row_index = header_row_index

    def __repr__(self):
        return f"Spreadsheet(product_spreadsheets='{self.product_spreadsheets}', product_sheet_name='{self.product_sheet_name}', headers={self.headers})"


class PaymentSheetModel:
    def __init__(self, payment_spreadsheet: str, payment_sheet_name: str, header_row_index: int):
        self.payment_spreadsheet = payment_spreadsheet
        self.payment_sheet_name = payment_sheet_name
        self.header_row_index = header_row_index

    def __repr__(self):
        return f"Spreadsheet(payment_spreadsheet='{self.payment_spreadsheet}', payment_sheet_name='{self.payment_sheet_name}', headers={self.headers})"