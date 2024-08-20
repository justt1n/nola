import glob
import os

import gspread
from google.oauth2 import service_account

from app.models.ProcessModel import FilterProcessModel
from app.models.ProductSheetModel import ProductSheetModel


def build_filer_process_from_file(path: str):
    if not os.path.exists(path):
        return
    txt = ""
    with open(path, 'r', encoding='utf-8') as file:
        txt = file.read()
    lines = txt.strip().split('\n')

    # First two lines are for payment spreadsheet and sheet name
    payment_spreadsheet = lines[0].split(":", 1)[1].strip().split("/")[5]
    payment_sheet_name = lines[1].split(":", 1)[1].strip()

    product_spreadsheets = {}

    # Remaining lines should come in groups of three
    index = 1
    for i in range(2, len(lines), 3):
        product_spreadsheet = lines[i].split(":", 1)[1].strip().split("/")[5]
        product_sheet_name = lines[i + 1].split(":", 1)[1].strip()
        headers = [header.strip() for header in lines[i + 2].split(":", 1)[1].split(",")]

        # Use the index as the key
        product_spreadsheets[str(index)] = ProductSheetModel(product_spreadsheet, product_sheet_name, headers)
        index += 1

    return FilterProcessModel(product_spreadsheets, payment_spreadsheet, payment_sheet_name)


def indices_to_cell(indices):
    row, col = indices
    col_str = ""
    while col >= 0:
        col_str = chr(col % 26 + ord('A')) + col_str
        col = col // 26 - 1
    return f"{col_str}{row + 1}"


def col_to_index(col_str):
    index = 0
    for char in col_str:
        index = index * 26 + ord(char.upper()) - ord('A') + 1
    return index - 1


def get_google_credentials():
    creds = None
    credentials = 'credentials.json'
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive',
    ]
    if os.path.exists(credentials):
        creds = service_account.Credentials.from_service_account_file(credentials)
    scoped_creds = creds.with_scopes(scopes)
    return scoped_creds


class FilterProcessService:

    def __init__(self):
        creds = get_google_credentials()
        self.gc = gspread.authorize(creds)

    def detect_ranges(self, spreadsheet_id, sheet_id):
        spreadsheet = self.gc.open_by_key(spreadsheet_id)
        sheet = spreadsheet.get_worksheet_by_id(sheet_id)

        max_rows = sheet.row_count
        max_cols = sheet.col_count

        values = sheet.get_all_values()

        non_empty_cells = set()
        for r_idx, row in enumerate(values):
            for c_idx, cell in enumerate(row):
                if cell:
                    non_empty_cells.add((r_idx, c_idx))

        visited = set()

        def find_range(r, c):
            if (r, c) in non_empty_cells and (r, c) not in visited:
                min_r, max_r = r, r
                min_c, max_c = c, c
                # Expand vertically
                for i in range(r, max_rows):
                    if (i, c) in non_empty_cells:
                        max_r = i
                    else:
                        break
                # Expand horizontally
                for j in range(c, max_cols):
                    if (r, j) in non_empty_cells:
                        max_c = j
                    else:
                        break
                # Mark all cells in the range as visited
                for i in range(min_r, max_r + 1):
                    for j in range(min_c, max_c + 1):
                        visited.add((i, j))
                return (min_r, min_c, max_r, max_c)
            return None

        ranges = []
        for r_idx in range(max_rows):
            for c_idx in range(max_cols):
                range_coords = find_range(r_idx, c_idx)
                if range_coords:
                    min_r, min_c, max_r, max_c = range_coords
                    range_str = f"{sheet.title}!{gspread.utils.rowcol_to_a1(min_r + 1, min_c + 1)}:{gspread.utils.rowcol_to_a1(max_r + 1, max_c + 1)}"
                    ranges.append(range_str)

        return ranges

    def filter_and_transfer_data(self, src_spreadsheets, des_spreadsheet_id, des_sheet_name):
        print("Start filtering and transferring data")
        des_spreadsheet = self.gc.open_by_key(des_spreadsheet_id)
        des_sheet = des_spreadsheet.worksheet(des_sheet_name)
        des_sheet.clear()
        col_offset = 0  # Column offset between data ranges

        for index, spreadsheet_info in src_spreadsheets.items():
            product_spreadsheet = self.gc.open_by_key(spreadsheet_info.product_spreadsheets)

            product_sheet = product_spreadsheet.worksheet(spreadsheet_info.product_sheet_name)
            values = product_sheet.get_all_values()
            values = [row for row in values if any(cell.strip() for cell in row)]
            if not values:
                continue

            # Detect the starting row and column using detect_ranges
            detected_ranges = self.detect_ranges(spreadsheet_info.product_spreadsheets, product_sheet.id)
            if not detected_ranges:
                continue

            # Get the starting cell of the first (and only) range
            start_cell, end_cell = detected_ranges[0].split(":")
            start_cell = start_cell.split("!")[1]
            start_row, start_col = gspread.utils.a1_to_rowcol(start_cell)

            # Extract header row from the detected start_row
            header_row = values[3]

            try:
                # Ensure all columns exist in the header row
                col_indices = [header_row.index(col_name) for col_name in spreadsheet_info.headers]
            except ValueError as e:
                raise ValueError(
                    f"Column '{e.args[0]}' not found in the header row of sheet '{spreadsheet_info.product_sheet_name}'")

            # Create the new header row for the filtered values
            new_header_row = [header_row[idx] for idx in col_indices] + ["Identifier"]
            status_col_index = header_row.index("Trạng thái")

            filtered_values = [new_header_row]  # Insert header at the beginning
            for row_idx, row in enumerate(values[1:], start=start_row):
                row = [item.strip().lower() for item in row]
                if 'unpaid' in row:
                    filtered_row = [row[idx] for idx in col_indices]
                    # Calculate begin_col and end_col based on the current col_offset
                    begin_col = gspread.utils.rowcol_to_a1(1, col_offset + 2)
                    end_col = gspread.utils.rowcol_to_a1(1, col_offset + 1 + len(spreadsheet_info.headers) + 1)
                    # Add identifier information into a cell separated by "#"
                    identifier = f"{spreadsheet_info.product_spreadsheets}#{spreadsheet_info.product_sheet_name}#{begin_col}:{end_col}#{indices_to_cell((row_idx, status_col_index))}"
                    filtered_row.append(identifier)
                    filtered_values.append(filtered_row)

            if filtered_values:
                start_col_offset = col_offset + 2  # Adjusted offset to 2
                end_col = start_col_offset + len(spreadsheet_info.headers) + 1  # +1 for the identifier information
                if end_col > des_sheet.col_count:
                    raise ValueError(
                        f"End column {end_col} exceeds the sheet's column count {des_sheet.col_count}")

                start_cell = gspread.utils.rowcol_to_a1(1, start_col_offset)
                end_cell = gspread.utils.rowcol_to_a1(len(filtered_values), end_col)
                cell_range = f"{start_cell}:{end_cell}"

                des_sheet.update(cell_range, filtered_values)
                col_offset += len(spreadsheet_info.headers) + 1 + 2  # Horizontal spacing between datasets

                # Set the identifier column to wrap text to clip
                identifier_col = start_col_offset + len(spreadsheet_info.headers)
                identifier_range = f"{gspread.utils.rowcol_to_a1(1, identifier_col)}:{gspread.utils.rowcol_to_a1(len(filtered_values), identifier_col)}"
                des_sheet.format(identifier_range, {"wrapStrategy": "CLIP"})
        return True

    def get_all_sheets(self, spreadsheet_id):
        spreadsheet = self.gc.open_by_key(spreadsheet_id)
        sheet_titles = []

        for sheet in spreadsheet.worksheets():
            sheet_titles.append(sheet.title)  # Store the title of each sheet

        return sheet_titles

    def get_header(self, spreadsheet_id, sheet_name):
        spreadsheet = self.gc.open_by_key(spreadsheet_id)
        sheet = spreadsheet.worksheet(sheet_name)
        values = sheet.get_all_values()
        values = [row for row in values if any(cell.strip() for cell in row)]
        start_row, start_col = 0, 0
        for i, row in enumerate(values):
            if any(cell.strip() for cell in row):
                start_row = i
                start_col = next((idx for idx, cell in enumerate(row) if cell.strip()), 0)
                break

        header_row = values[start_row]
        header_row = [col for col in header_row if col]
        return header_row

    def format_status_column(self, spreadsheet_id, sheet_id, range):
        spreadsheet = self.gc.open_by_key(spreadsheet_id)
        sheet = spreadsheet.get_worksheet_by_id(sheet_id)

        header_row = 3  # assuming the first three rows are notes
        headers = sheet.row_values(header_row)
        status_col = headers.index("Trạng thái") + 1  # gspread is 1-indexed

        # Set data validation (dropdown list) in the "Trạng thái" column
        validation_rule = {
            "requests": [
                {
                    "setDataValidation": {
                        "range": {
                            "sheetId": sheet.id,
                            "startRowIndex": header_row,  # start from header_row (0-indexed)
                            "endRowIndex": sheet.row_count,  # until the last row
                            "startColumnIndex": status_col - 1,  # 0-indexed
                            "endColumnIndex": status_col  # 0-indexed
                        },
                        "rule": {
                            "condition": {
                                "type": "ONE_OF_LIST",
                                "values": [
                                    {"userEnteredValue": "paid"},
                                    {"userEnteredValue": "unpaid"}
                                ]
                            },
                            "showCustomUi": True
                        }
                    }
                }
            ]
        }

        # Apply the validation rule to the sheet
        sheet.batch_update(validation_rule)

    def processSingle(self, path):
        # load all payload from files
        payload = build_filer_process_from_file(path)

        # process the payload
        response = self.filter_and_transfer_data(payload.product_spreadsheets, payload.payment_spreadsheet,
                                                 payload.payment_sheet_name)
        print("Filter Completed")
        if response:
            # move the file to done folder
            done_filename = os.path.join('done', os.path.basename(path))
            os.rename(path, done_filename)
            print("Done")
            return True
        print("Failed")
        return False

    def processMultiple(self):
        # load all payload files from pending folder
        pending_files = glob.glob("pending/*.txt")
        for path in pending_files:
            self.processSingle(path)
