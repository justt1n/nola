from typing import Dict

from app.models.SheetModel import PaymentSheetModel
from app.models.SheetModel import ProductSheetModel


class FilterProcessModel:
    def __init__(self, product_spreadsheets: Dict[str, ProductSheetModel], payment_spreadsheet: PaymentSheetModel):
        self.product_spreadsheets = product_spreadsheets
        self.payment_spreadsheet = payment_spreadsheet

    def __repr__(self):
        return f"MultiSpreadsheetFilter(Products={self.product_spreadsheets}, Payment='{self.payment_spreadsheet}')"
