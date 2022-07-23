from .types import PasteType
import re


class ActionChain(object):
    AUTO = -1

    def __init__(self, spreadsheet):
        self._spreadsheet = spreadsheet
        self._requests = []

    def update_cells(self, sheet, name_box, data):
        pass

    def cut_paste(self, sheet: str, name_box_cut: str, name_box_paste: str, paste_type=PasteType.FORMULA) -> None:
        _id = self._spreadsheet.sheets[sheet]["sheetId"]
        row = self._spreadsheet.sheets[sheet]["rowCount"]
        col = self._spreadsheet.sheets[sheet]["columnCount"]
        cut = self.index_name_box(name_box_cut, row, col)
        cut["sheetId"] = _id
        paste = self.index_name_box(name_box_paste, row, col)
        self._requests.append({
            "cutPaste": {
                "source": cut,
                "destination": {
                    "sheetId": _id,
                    "rowIndex": paste["startRowIndex"],
                    "columnIndex": paste["startColumnIndex"]
                },
                "pasteType": paste_type
            }
        })

    def copy_paste(self, sheet: str, name_box_copy: str, name_box_paste: str, paste_type=PasteType.FORMULA) -> None:
        _id = self._spreadsheet.sheets[sheet]["sheetId"]
        row = self._spreadsheet.sheets[sheet]["rowCount"]
        col = self._spreadsheet.sheets[sheet]["columnCount"]
        copy = self.index_name_box(name_box_copy, row, col)
        paste = self.index_name_box(name_box_paste, row, col)
        copy["sheetId"] = _id
        paste["sheetId"] = _id
        self._requests.append({
            "copyPaste": {
                "source": copy,
                "destination": paste,
                "pasteType": paste_type
            }
        })

    def duplicate(self, sheet: str, _as: str) -> None:
        self._requests.append({
            "duplicateSheet": {
                "sourceSheetId": self._sheets[sheet]["sheetId"],
                "newSheetName": _as
            }
        })

    def resize_row(self, sheet: str, name_box: str, pixel_size: int = AUTO) -> None:
        _id = self._spreadsheet.sheets[sheet]["sheetId"]
        row = self._spreadsheet.sheets[sheet]["rowCount"]
        col = self._spreadsheet.sheets[sheet]["columnCount"]
        index = self.index_name_box(name_box, row, col, self.ROWS)
        if pixel_size == self.AUTO:
            self._requests.append({
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": _id,
                        "dimension": "ROWS",
                        "startIndex": index["startRowIndex"],
                        "endIndex": index["endRowIndex"]
                    }
                }
            })
        else:
            assert 2 <= pixel_size <= 2000, "Pixel size must be between 2 and 2000"
            self._requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": _id,
                        "dimension": "ROWS",
                        "startIndex": index["startRowIndex"],
                        "endIndex": index["endRowIndex"]
                    }
                }
            })

    def resize_column(self, sheet: str, name_box: str, pixel_size: int = AUTO) -> None:
        _id = self._spreadsheet.sheets[sheet]["sheetId"]
        row = self._spreadsheet.sheets[sheet]["rowCount"]
        col = self._spreadsheet.sheets[sheet]["columnCount"]
        index = self.index_name_box(name_box, row, col, self.COLUMNS)
        if pixel_size == self.AUTO:
            self._requests.append({
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": _id,
                        "dimension": "COLUMNS",
                        "startIndex": index["startColumnIndex"],
                        "endIndex": index["endColumnIndex"]
                    }
                }
            })
        else:
            assert 2 <= pixel_size <= 2000, "Pixel size must be between 2 and 2000"
            self._requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": _id,
                        "dimension": "COLUMNS",
                        "startIndex": index["startColumnIndex"],
                        "endIndex": index["endColumnIndex"]
                    }
                }
            })

    def append_row(self, sheet: str, number: int) -> None:
        self._requests.append({
            "appendDimension": {
                "sheetId": self._spreadsheet.sheets[sheet]["sheetId"],
                "dimension": "ROWS",
                "length": number
            }
        })

    def append_column(self, sheet: str, number: int) -> None:
        self._requests.append({
            "appendDimension": {
                "sheetId": self._spreadsheet.sheets[sheet]["sheetId"],
                "dimension": "COLUMNS",
                "length": number
            }
        })

    def delete_row(self, sheet: str, name_box: str) -> None:
        row = self._spreadsheet.sheets[sheet]["rowCount"]
        col = self._spreadsheet.sheets[sheet]["columnCount"]
        index = self.index_name_box(name_box, row, col, self.ROWS)
        self._requests.append({
            "deleteDimension": {
                "range": {
                    "sheetId": self._spreadsheet.sheets[sheet]["sheetId"],
                    "dimension": "ROWS",
                    "startIndex": index["startRowIndex"],
                    "endIndex": index["endRowIndex"]
                }
            }
        })

    def delete_column(self, sheet: str, name_box: str) -> None:
        row = self._spreadsheet.sheets[sheet]["rowCount"]
        col = self._spreadsheet.sheets[sheet]["columnCount"]
        index = self.index_name_box(name_box, row, col, self.COLUMNS)
        self._requests.append({
            "deleteDimension": {
                "range": {
                    "sheetId": self._spreadsheet.sheets[sheet]["sheetId"],
                    "dimension": "COLUMNS",
                    "startIndex": index["startColumnIndex"],
                    "endIndex": index["endColumnIndex"]
                }
            }
        })

    def insert_row(self, sheet: str, number: int, start_row: int) -> None:
        self._requests.append({
            "insertDimension": {
                "range": {
                    "sheetId": self._spreadsheet.sheets[sheet]["sheetId"],
                    "dimension": "ROWS",
                    "startIndex": start_row - 1,
                    "endIndex": start_row + number
                }
            }
        })

    def insert_column(self, sheet: str, number: int, start_column: str) -> None:
        col = self.str2count(start_column)
        self._requests.append({
            "insertDimension": {
                "range": {
                    "sheetId": self._spreadsheet.sheets[sheet]["sheetId"],
                    "dimension": "COLUMNS",
                    "startIndex": col - 1,
                    "endIndex": col + number
                }
            }
        })

    def perform(self) -> dict:
        response = self._spreadsheet.execute({"requests": self._requests})
        self._requests.clear()
        return response

    @staticmethod
    def str2count(string: str) -> int:
        return sum(26**i * (ord(s) - 64) for i, s in enumerate(list(string)[::-1]))

    @staticmethod
    def count2str(count: int) -> str:
        chars, i, c = "", 0, count
        while sum(26**a for a in range(i + 2)) <= c:
            i += 1
        for k in range(i, 0, -1):
            x, j = 0, 0
            while j < 26:
                j += 1
                if 26**k * j >= c:
                    j -= 1
                    break
                x = 26**k * j
            c -= x
            chars += chr(j + 64)
        return chars + chr(c + 64)

    ROWS = "ROWS"
    COLUMNS = "COLUMNS"
    NORMAL = "NORMAL"
    @staticmethod
    def index_name_box(name_box: str, row_count: int, column_count: int, _format: str = "NORMAL") -> dict:
        if _format == "ROWS":
            if not re.search(r"^\d+:\d+$", name_box):
                raise ValueError("Invalid name box: %s" % name_box)
        elif _format == "COLUMNS":
            if not re.search(r"^[A-Z]+:[A-Z]+$", name_box):
                raise ValueError("Invalid name box: %s" % name_box)
        else:
            if not re.search(r"^[A-Z]+\d+:[A-Z]+\d+$|^\d+:\d+$|^[A-Z]+:[A-Z]+$|^[A-Z]+\d+$", name_box):
                raise ValueError("Invalid name box: %r" % name_box)
        n = name_box.split(":")
        start_col = re.search(r"^[A-Z]+", n[0])
        if start_col is None:
            start_col_index = 0
            start_row_index = int(n[0]) - 1
        else:
            start_col_index = ActionChain.str2count(start_col.group()) - 1
            start_row = re.sub(start_col.group(), "", n[0])
            if not start_row:
                start_row_index = 0
            else:
                start_row_index = int(start_row) - 1
        if len(n) == 2:
            end_col = re.search(r"^[A-Z]+", n[1])
            if end_col is None:
                end_col_index = column_count
                end_row_index = int(n[1])
            else:
                end_col_index = ActionChain.str2count(end_col.group())
                end_row = re.sub(end_col.group(), "", n[1])
                if not end_row:
                    end_row_index = row_count
                else:
                    end_row_index = int(end_row)
        else:
            end_col_index = start_col_index + 1
            end_row_index = start_row_index + 1
        return {
            "startRowIndex": start_row_index,
            "endRowIndex": end_row_index,
            "startColumnIndex": start_col_index,
            "endColumnIndex": end_col_index
        }
