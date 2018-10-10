import httplib2
from urllib import request
from bs4 import BeautifulSoup
from openpyxl import load_workbook

http = httplib2.Http()
KEYWORDS = ['Downlink', 'Uplink', 'Offset', 
        'Uplink Tone', 'Downlink Tone', 'Call', 
        'Use', 'Sponsor', 'Affaliate', 
        'Links', 'EchoLink', 'IRPL', 
        'Last update']

def load_rb(cell):
    status, response = http.request(cell.hyperlink.target)
    soup = BeautifulSoup(response, 'html.parser')

    details = {}

    for x in soup.find_all('td'):
        rb_cell = str(x.renderContents().strip())

        for keyword in KEYWORDS:
            if rb_cell.find(keyword) != -1 and x.find_next_sibling():
                details[keyword] = x.find_next_sibling().renderContents().strip()


    print(cell.value)
    print(details)

def main():
    workbook = load_workbook('radio.xlsx')
    sheet = workbook['Repeaters']
    row_num = 2
    col_name = 'A'
    cell_pos = "%s%d" % (col_name, row_num)
    cell = sheet[cell_pos]

    while cell.value is not None:
        if cell.hyperlink:
            load_rb(cell)

        row_num += 1
        cell_pos = "%s%d" % (col_name, row_num)
        cell = sheet[cell_pos]


main()
