import pygsheets

client = pygsheets.authorize(client_secret='oauth.json')
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1CIpWBZXAK_PqsEq2vLlFklT2MgQjH008qbxgPYXbRtU/edit#gid=0")
worksheets = spreadsheet.worksheets()

# Cell columns
orig_cells = spreadsheet.sheet1.get_col(1, include_tailing_empty=False, returnas='cells')
phone_cells = spreadsheet.sheet1.get_col(2, include_tailing_empty=False, returnas='cells')
web_cells = spreadsheet.sheet1.get_col(3, include_tailing_empty=False, returnas='cells')
add_cells = spreadsheet.sheet1.get_col(4, include_tailing_empty=False, returnas='cells')

# Declare lists and matrix
orig_list = list()
phone_list = list()
web_list = list()
add_list = list()
data_matrix = [[]]
data_matrix.clear()

# Get lists from columns
def list_function(cells, list):
    for c in cells:
        v = c.value
        list.append(v)

list_function(orig_cells, orig_list)
list_function(phone_cells, phone_list)
list_function(web_cells, web_list)
list_function(add_cells, add_list)

# Remove phone numbers, websites, and addresses from list items
for i, l in enumerate(orig_list):

    try:
        if phone_list[i] in l:
            l = l.replace(phone_list[i], "")
    except: pass

    try:
        if web_list[i] in l:
            l = l.replace(web_list[i], "")
    except: pass

    try:
        if add_list[i] in l:
            l = l.replace(add_list[i], "")
    except: pass

    data_matrix.append([l])

spreadsheet.sheet1.update_values("A:A", data_matrix)
