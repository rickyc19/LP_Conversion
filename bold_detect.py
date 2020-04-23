import pygsheets
import phonenumbers
from urlextract import URLExtract
import pyap
import os

client = pygsheets.authorize(client_secret='oauth.json')
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1CIpWBZXAK_PqsEq2vLlFklT2MgQjH008qbxgPYXbRtU/edit#gid=0")
sheet = spreadsheet.sheet1

cells = sheet.get_col(1, include_tailing_empty=False, returnas='cells')
value_list = list()
cell_list = list()
cell_ranges = list()

for c in cells:
    v = c.value

    if c.color == (0.8509804, 0.91764706, 0.827451, 0):
        value_list.append(v)
        cell_list.append(c.label)

print(value_list)
print(cell_list)

for i, l in enumerate(cell_list):
    try: cell_ranges.append(l + ":" + cell_list[i+1])
    except: pass

#cell_ranges = [cell_list[i] + ':' + cell_list[i+1] for i in cell_list]

print(cell_ranges)
