import httplib2
import string
from urllib import request
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, colors, Color

http = httplib2.Http()

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
def _AS_TEXT(value): return str(value) if value is not None else ""


class RepeaterBot:
    rb_url = 'https://www.repeaterbook.com'
    wb_dest = 'repeaters.xlsx'

    def __init__(self, callsigns=None):
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = 'Repeaters'
        self.callsigns = callsigns

        # build headers
        self.ws['A1'].font = _FONT_HEADER
        self.ws['A1'] = 'Callsign'

        for word in _KEYWORDS:
            coords = "%s1" % _LETTERS[_KEYWORDS.index(word)+1]
            self.ws[coords].font = _FONT_HEADER
            self.ws[coords] = word[0]

    def search_rb(self, callsign):
        status, response = http.request("%s/repeaters/callResult.php?call=%s&submit=RepeaterBook" % (self.rb_url, callsign))
        soup = BeautifulSoup(response, 'html.parser')

        links = []

        for x in soup.find_all('a', title='View details'):
            links.append(x['href'])

        return links

    def load_rb(self, rb_href):
        status, response = http.request("%s/repeaters/%s" % (self.rb_url, rb_href))
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
    # load callsigns from input file
    with open('repeaters.txt') as f:
        callsigns = f.readlines()

    # start repeaterbot and pass callsigns
    rpb = RepeaterBot([x.strip() for x in callsigns])

    # populate rows
    row = 2
    for sign in rpb.callsigns:
        for url in rpb.search_rb(sign):
            repeater_data = rpb.load_rb(url)
            rpb.ws["A%d" % (row)] = sign
            rpb.ws["A%d" % (row)].font = _FONT_BODY

            for word in _KEYWORDS:
                col = _LETTERS[_KEYWORDS.index(word)+1]
                coords = "%s%d" % (col, row)

                rpb.ws[coords].font = _FONT_BODY
                rpb.ws[coords] = repeater_data[word[1]]

            row += 1

    # fix widths
    for column_cells in rpb.ws.iter_cols(max_col=len(_KEYWORDS), max_row=row):
            length = max(len(_AS_TEXT(cell.value)) for cell in column_cells) * 1.5

            if length > _MIN_COL_WIDTH:
                rpb.ws.column_dimensions[column_cells[0].column].width = length
            else:
                rpb.ws.column_dimensions[column_cells[0].column].width = _MIN_COL_WIDTH

    rpb.wb.save(filename = rpb.wb_dest)


if __name__ == "__main__" :
    main()
