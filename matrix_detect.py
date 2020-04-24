import phonenumbers
from urlextract import URLExtract
import pyap

def find_info(value_matrix, sheet):
    phone_string = ""
    data_matrix = [[]]
    data_matrix.clear()

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
        phone_string = ""

    sheet.update_values("C:H", data_matrix)  # Update cell range with found values

