from openpyxl.styles import Border, Side, PatternFill
from openpyxl.worksheet.worksheet import Worksheet
from unicodedata import normalize

thin_border = Border(
    left = Side(style = 'thin'), 
    right = Side(style = 'thin'), 
    top = Side(style = 'thin'), 
    bottom = Side(style = 'thin')
)

graybg = PatternFill(start_color = 'cccccc', fill_type = 'solid')

def remove_acc(word: str):
    norm_word = normalize('NFKD', word).encode('ASCII', 'ignore').decode('ASCII')
    return norm_word

def create_filter(value: str, num: bool = False):
    temp_filters = remove_acc(value.upper()).split('|')
    filters = map(lambda x: x.strip(), temp_filters)

    if num:
        return list(map(lambda x: int(x), filters))

    return list(filters)

def apply_filter(descr: str, words: list[str], func):
    if words:
        return func(x in descr for x in words)
    else:
        return True

def color_row(row: tuple, color: str):
    background = PatternFill(start_color = color, fill_type = 'solid')
    for cell in row:
        cell.fill = background
        cell.border = thin_border

def copy_sheet(ws_source: Worksheet, ws_destination: Worksheet):
    mr = ws_source.max_row
    mc = ws_source.max_column
    for i in range (1, mr + 1):
        for j in range (1, mc + 1):
            c = ws_source.cell(row = i, column = j)
            ws_destination.cell(row = i, column = j).value = c.value