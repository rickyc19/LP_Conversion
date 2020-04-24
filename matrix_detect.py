import pygsheets
import phonenumbers
from urlextract import URLExtract
import pyap
import os

client = pygsheets.authorize(client_secret='oauth.json')
spreadsheet = client.open_by_url(
    "https://docs.google.com/spreadsheets/d/1CIpWBZXAK_PqsEq2vLlFklT2MgQjH008qbxgPYXbRtU/edit#gid=0")
sheet = spreadsheet.sheet1

provider_string = ""
phone_string = ""
web_string = ""
add_string = ""

value_list = list()
data_matrix = [[]]
data_matrix.clear()

# Dummy matrix for testing
value_matrix = [['Edgewood Vista', '4195 Westberg Road, Hermantown MN 55811, (218) 723-8905, www.edgewoodvista.com'],
                ['Floodwood Adult Day Services',
                 '601 Ash Street, PO Box 347, Floodwood MN 55736, (218) 476-2230 Donna Tracy, Stacy, Stevens, www.fst2b.org']]

for v in value_matrix:

    # Provider name from matrix
    provider_string = v[0]

    try:
        add_string = pyap.parse(v[1], country='US')[0].__str__()  # Find address from matrix value index 1
    except:
        add_string = ""  # Expect error when no address found. Make value "" to add to matrix

    try:
        web_string = URLExtract().find_urls(v[1])[0]  # Find URL from matrix value index 1
    except:
        web_string = ""  # Expect error when no URL found. Make value "" to add to matrix

    for match in phonenumbers.PhoneNumberMatcher(v[1], "US"):  # Find phone number from matrix value index 1
        phone_string = phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.NATIONAL)

    data_matrix.append([provider_string, phone_string, web_string, add_string])

sheet.update_values("A:D", data_matrix)  # Update cell range with found values
