import httplib2
import string
from urllib import request
from bs4 import BeautifulSoup
from openpyxl import Workbook

http = httplib2.Http()
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

    for x in soup.find_all('td'):
        rb_cell = str(x.renderContents().strip())

        for keyword in _KEYWORDS:
            if rb_cell.find(keyword[0]) != -1 and x.find_next_sibling():
                details[keyword[1]] = x.find_next_sibling().renderContents().strip()

    return details

def main():
    wb = Workbook()
    wb_dest = 'repeaters.xlsx'
    ws = wb.active
    ws.title = 'Repeaters'

    with open('repeaters.txt') as f:
        callsigns = f.readlines()

    callsigns = [x.strip() for x in callsigns]
    row_num = 2

    _LETTERS = string.ascii_uppercase

    for word in _KEYWORDS:
        ws["%s1" % _LETTERS[_KEYWORDS.index(word)]] = word[0]

    for sign in callsigns:
        for url in search_rb(sign):
            repeater_data = load_rb(url)
            ws["A%d" % (row_num)] = sign

            for word in _KEYWORDS:
                print(hasattr(repeater_data, word[1]))
                print(word[1])
                print(repeater_data[word[1]])
                # if hasattr(repeater_data, word[1]):
                    # print(repeater_data[word[1]])
                    # ws["%s%d" % (_LETTERS[_KEYWORDS.index(word)], row_num)] = repeater_data[word[1]]

            row_num += 1

    wb.save(filename = wb_dest)


main()
