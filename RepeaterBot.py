import httplib2
from urllib import request
from bs4 import BeautifulSoup
from openpyxl import load_workbook

http = httplib2.Http()

def search_rb(callsign):
    status, response = http.request("https://www.repeaterbook.com/repeaters/callResult.php?call=%s&submit=RepeaterBook" % callsign)
    soup = BeautifulSoup(response, 'html.parser')

    links = []

    for x in soup.find_all('a', title='View details'):
        links.append(x['href'])

    return links

def load_rb(rb_href):
    KEYWORDS = ['Downlink', 'Uplink', 'Offset', 
            'Uplink Tone', 'Downlink Tone', 'Call', 
            'Use', 'Sponsor', 'Affaliate', 
            'Links', 'EchoLink', 'IRPL', 
            'Last update']

    status, response = http.request("https://www.repeaterbook.com/repeaters/%s" % rb_href)
    soup = BeautifulSoup(response, 'html.parser')

    details = {}

    for x in soup.find_all('td'):
        rb_cell = str(x.renderContents().strip())

        for keyword in KEYWORDS:
            if rb_cell.find(keyword) != -1 and x.find_next_sibling():
                details[keyword] = x.find_next_sibling().renderContents().strip()

    return details

def main():
    workbook = load_workbook('radio.xlsx')
    sheet = workbook['Repeaters']
    row_num = 2
    col_name = 'A'
    cell_pos = "%s%d" % (col_name, row_num)
    cell = sheet[cell_pos]

    repeaters = {}

    while cell.value is not None:
        repeaters[cell.value] = { "links": search_rb(cell.value)}

        row_num += 1
        cell_pos = "%s%d" % (col_name, row_num)
        cell = sheet[cell_pos]

    for callsign in repeaters:
        for link in repeaters[callsign]['links']:
            print(load_rb(link))


main()
