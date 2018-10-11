import httplib2
import string
from urllib import request
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, colors, Color

http = httplib2.Http()
def _AS_TEXT(value): return str(value) if value is not None else ""
_MIN_COL_WIDTH = 30
_LETTERS = string.ascii_uppercase
_KEYWORDS = [('Downlink', 'downlink'), 
        ('Uplink', 'uplink'),
        ('Offset', 'offset'),
        ('Uplink Tone', 'uplinktone'),
        ('Downlink Tone', 'downlinktone'),
        ('Call', 'call'),
        ('Use', 'use'),
        ('Sponsor', 'sponsor'),
        ('Affaliate', 'affaliate'),
        ('Links', 'links'),
        ('EchoLink', 'echolink'),
        ('IRPL', 'irpl'),
        ('Last update','lastupdate')]
_FONT_HEADER = Font(name='Calibri',
        color=colors.BLUE,
        size=18)
_FONT_BODY = Font(name='Calibri',
        size=14)


def search_rb(callsign):
    status, response = http.request("https://www.repeaterbook.com/repeaters/callResult.php?call=%s&submit=RepeaterBook" % callsign)
    soup = BeautifulSoup(response, 'html.parser')

    links = []

    for x in soup.find_all('a', title='View details'):
        links.append(x['href'])

    return links

def load_rb(rb_href):
    status, response = http.request("https://www.repeaterbook.com/repeaters/%s" % rb_href)
    soup = BeautifulSoup(response, 'html.parser')

    details = {}
    for keyword in _KEYWORDS:
        details[keyword[1]] = None

    for x in soup.find_all('td'):
        rb_cell = str(x.renderContents().strip())

        for keyword in _KEYWORDS:
            cell_title = "%s:" % keyword[0]
            if rb_cell.find(cell_title) != -1 and x.find_next_sibling():
                details[keyword[1]] = x.find_next_sibling().renderContents().strip()

    return details

def main():
    wb = Workbook()
    wb_dest = 'repeaters.xlsx'
    ws = wb.active
    ws.title = 'Repeaters'

    # load callsigns from input file
    with open('repeaters.txt') as f:
        callsigns = f.readlines()

    # build headers
    for word in _KEYWORDS:
        coords = "%s1" % _LETTERS[_KEYWORDS.index(word)]
        ws[coords].font = _FONT_HEADER
        ws[coords] = word[0]

    callsigns = [x.strip() for x in callsigns]
    row = 2

    # populate rows
    for sign in callsigns:
        for url in search_rb(sign):
            repeater_data = load_rb(url)
            ws["A%d" % (row)] = sign

            for word in _KEYWORDS:
                col = _LETTERS[_KEYWORDS.index(word)]
                coords = "%s%d" % (col, row)

                ws[coords].font = _FONT_BODY
                ws[coords] = repeater_data[word[1]]

            row += 1

    # fix widths
    for column_cells in ws.iter_cols(max_col=len(_KEYWORDS), max_row=row):
            length = max(len(_AS_TEXT(cell.value)) for cell in column_cells) * 1.5

            if length > _MIN_COL_WIDTH:
                ws.column_dimensions[column_cells[0].column].width = length
            else:
                ws.column_dimensions[column_cells[0].column].width = _MIN_COL_WIDTH

    wb.save(filename = wb_dest)


if __name__ == "__main__" :
    main()
