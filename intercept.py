import xlrd


def excel_file(filename):
    if filename.endswith(('.xlsx', '.xls', '.csv')):
        # open excel and sheet
        workbook = xlrd.open_workbook(filename)
        sheet = workbook.sheet_by_index(0)

        # Stores Position and Partcode value pairs from all rows.
        all_data = []

        # loop over rows and columns to extract Position and Partcode pairs
        for row in range(sheet.nrows):
            current_row = []
            for col in range(sheet.ncols):
                if col == 1 or col == 3:
                    data = sheet.cell_value(row, col)
                    # useful when searching in CC
                    new_data = data.replace('/', "_")
                    current_row.append(new_data)
            all_data.append(current_row)

        print(all_data)

    else:
        print(f'{filename} is not excel file.')


if __name__ == "__main__":
    excel_file(filename='excel.xlsx')
