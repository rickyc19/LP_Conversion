# Every row is added to both columns of matrix
def every_row(cells):
    final_sheet = [['Program Name', 'Address 1', 'Address 2', 'Phone 1', 'Phone 2', 'Fax', 'URL', 'Email', 'Notes']]
    row = 0

    for c in cells:
        final_sheet.append([c.value, ' '])
        final_sheet[row][1] += c.value
        row += 1

    return final_sheet


# For each cell: bolded ones are a new row for final sheet, others are information for the 2nd cell of same row.
def bold(cells):
    final_sheet = [['Program Name', 'Address 1', 'Address 2', 'Phone 1', 'Phone 2', 'Fax', 'URL', 'Email', 'Notes']]
    row = 0

    for c in cells:

        cj = c.get_json()

        if 'bold' in cj['userEnteredFormat']['textFormat']:
            # if cj['userEnteredFormat']['textFormat']['fontSize'] == 21:
            if final_sheet[row][1]:
                final_sheet.append([c.value, ' '])
                row += 1
            # If we haven't added any information on the program yet (even a blank row) then this line must be more title
            else:
                final_sheet[row][0] += ' ' + c.value
            # TODO: Distinguish between consecutive single line programs and titles split over multiple lines
        else:
            if final_sheet[row][1]:
                final_sheet[row][1] += ' '
            final_sheet[row][1] += c.value

    return final_sheet


