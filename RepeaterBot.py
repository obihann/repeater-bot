from openpyxl import load_workbook
workbook = load_workbook('radio.xlsx')
sheet = workbook['Repeaters']


row_num = 2
col_name = 'A'
cell_pos = "%s%d" % (col_name, row_num)
cell = sheet[cell_pos]

while cell.value is not None:
    print(cell.value)

    if cell.hyperlink:
        print(cell.hyperlink.target)

    row_num += 1
    cell_pos = "%s%d" % (col_name, row_num)
    cell = sheet[cell_pos]
