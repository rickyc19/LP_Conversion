import pygsheets
from color_detect import *
from matrix_detect import *

client = pygsheets.authorize(client_secret='OAuth.json')
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1CIpWBZXAK_PqsEq2vLlFklT2MgQjH008qbxgPYXbRtU/edit#gid=0")
sheet = spreadsheet.sheet1
cells = sheet.get_col(1, include_tailing_empty=False, returnas='cells')

# Create data matrix from cells
final_sheet = create_matrix(cells)
# Get necessary info from data matrix and put in correct columns
find_info(final_sheet, sheet)

