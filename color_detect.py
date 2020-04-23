import pygsheets
import phonenumbers
from urlextract import URLExtract
import pyap
import os

client = pygsheets.authorize(client_secret='oauth.json')
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1CIpWBZXAK_PqsEq2vLlFklT2MgQjH008qbxgPYXbRtU/edit#gid=0")
sheet = spreadsheet.sheet1

cells = sheet.get_col(1, include_tailing_empty=False, returnas='cells')

# Create the final sheet we will eventually update back. Headers are just an example at this point.
final_sheet = [['Program Name:', 'Address 1', 'Address 2', 'Phone 1', 'Phone 2', 'Fax', 'URL', 'Email', 'Notes']]

# Remove cells from list until we get to the first highlighted cell, ie the first program
while cells[0].color != (0.8509804, 0.91764706, 0.827451, 0):
    cells.pop(0)

# Row tracks which row of the final sheet we are working on.
row = 0

# For each cell: highlighted ones are a new row for final sheet, others are information for the 2nd cell of same row.
for c in cells:
    if c.color == (0.8509804, 0.91764706, 0.827451, 0):
        final_sheet.append([c.value, ''])
        row += 1
    else:
        final_sheet[row][1] += c.value + ', '  # This adds a excess comma and space at the end of the string.
