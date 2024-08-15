from typing import Dict

from app.models.ProductSheetModel import ProductSheetModel


class FilterProcessModel:
    def __init__(self, product_spreadsheets: Dict[str, ProductSheetModel], payment_spreadsheet: str,
                 payment_sheet_name: str):
        self.product_spreadsheets = product_spreadsheets
        self.payment_spreadsheet = payment_spreadsheet
        self.payment_sheet_name = payment_sheet_name

    def __repr__(self):
        return f"MultiSpreadsheetFilter(product_spreadsheets={self.product_spreadsheets}, payment_spreadsheet='{self.payment_spreadsheet}', payment_sheet_name='{self.payment_sheet_name}')"
