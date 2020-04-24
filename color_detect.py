def create_matrix(cells):
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
            if final_sheet[row][1]:
                final_sheet.append([c.value, ''])
                row += 1
            # If we haven't added any information on the program yet (even a blank row) then this line must be more title
            else:
                final_sheet[row][0] += ' ' + c.value
            # TODO: Distinguish between consecutive single line programs and titles split over multiple lines
        else:
            if final_sheet[row][1]:
                final_sheet[row][1] += ', '
            final_sheet[row][1] += c.value

    return final_sheet

