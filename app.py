import openpyxl
import re
from openpyxl.worksheet.worksheet import Worksheet

workbook = openpyxl.load_workbook(filename="data.xlsx")
sheet = workbook.active
last_column = "AD"

def get_words(row):
    _words = []
    for item in row[9:]:
        _value = item.value
        if not _value:
            continue
        _word = re.findall(r"\w+", str(_value))
        if _word:
            _words.append(_value)
    return _words

def main(sheet: Worksheet):
    max_row = sheet.max_row
    for row_no in range(max_row, 1, -1):
        words = get_words(sheet[row_no])
        if not words:
            continue
        sheet.move_range(f"A{row_no + 1}:AD{max_row + 1}", rows=len(words), cols=0)
        for cell in sheet[row_no][9:]:
            cell.value = None
        for jdx, word in enumerate(words):
            sheet[f"I{row_no + jdx + 1}"] = word

main(sheet)
workbook.save("result.xlsx")