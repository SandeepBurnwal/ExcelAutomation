import os
import sys
import time
import win32api
import win32con
import numpy as np
import xlwings as xw


def main(path_of_file):

    """ Path of the project directory (of the excel workbook) is being passed from
        VBA RunPython module when the Python script is called """

    # Set the working directory as per the workbook path received from VBA (ThisWorkbook.Path)
    os.chdir(path_of_file)

    # connects to the workbook to be modified/ processed !!
    wb = xw.Book('my_math.xlsm')

    parent_function(wb)
#############################################################################


def parent_function(wb):
    """ This is the parent function which subsequently calls 2 downstream function. One of the function (write_data) writes the found + missing data back in the excell sheet
        Other downstream function (analytics) writes back data counts + plots. This would help in reconciliation """

    # sheet I would like to modify
    data_sheet = wb.sheets['Data']
    odd_number = int(data_sheet.range('A3').value)
    cell_number = int(data_sheet.range('B3').value)

    # Clear the old data
    data_sheet.range('D:Z').api.Delete()

    # Validate if the cell_number is an odd number and  pass back the message to the user accordingly
    odd = (cell_number % 2)
    if odd == 0:
        win32api.MessageBox(xw.apps.active.api.Hwnd, f'{cell_number} is even, pls enter odd no !! ', 'Info', win32con.MB_ICONINFORMATION)
        raise SystemExit(0)
    elif cell_number > 17:
        win32api.MessageBox(xw.apps.active.api.Hwnd, f'Pls enter a odd no less than 17 !! ', 'Info', win32con.MB_ICONINFORMATION)
        raise SystemExit(0)

    # Call the number calculation logic function
    calculate_numbers(data_sheet, odd_number, cell_number)

    # Passing back a text message to the user
    sum_number = int(odd_number / cell_number) * cell_number
    message = f'You can validate, the sum is always {sum_number} (nearest/exact multiple of {cell_number}) in any direction !!'
    # for i in range(1, len(message) + 1):
    for i, msg in enumerate(message, 1):
        data_sheet.range('D2:J2').color = (255, 255, 0)
        data_sheet.range('D2').api.Font.Bold = True
        data_sheet.range('D2').value = message[:i]
        time.sleep(0.010)
##############################################################################################


def calculate_numbers(data_sheet, odd_number, cell_number):
    # set the starting location in the matrix
    row = 0
    col = int(cell_number / 2)

    total_elements = cell_number * cell_number
    matrix = np.arange(total_elements).reshape(cell_number, cell_number)
    matrix[:, :] = 0

    # find the sequence count of the odd cell_number .. like 3 is 1st, 5 is 2nd, 7 is 3rd, 9 is 4th, 11 is 5th etc
    seq_count = int(cell_number / 2)
    subtract_no = 2*(seq_count*seq_count + seq_count)  # 2(r2+r) r is the count of the number

    next_number = int(odd_number / cell_number) - subtract_no  # find the starting number in the matrix
    count = 0
    while count < total_elements:
        matrix[row][col] = next_number
        count = count + 1
        next_number = next_number + 1
        row = row - 1
        col = col + 1

        row, col = matrix_edges_logic(cell_number, matrix, row, col)

    data_sheet.range('D5').value = matrix
    border(data_sheet, 'D5', 'F5')

    initial_range = data_sheet.range('D5')
    for i in range(0, cell_number):
        calculate_col_total(initial_range.offset(0, i), initial_range.offset(cell_number, i))
        calculate_row_total(initial_range.offset(i, 0), initial_range.offset(i, cell_number))

##############################################################################################


def matrix_edges_logic(cell_number, matrix, row, col):
    if ((row < cell_number) and (col < cell_number) and matrix[row][col] != 0) or ((row < 0) and (col == cell_number)):
        row = row + 2
        col = col - 1

    elif (row >= 0) and (col == cell_number):
        col = 0

    elif row == -1:
        row = cell_number - 1

    else:
        pass

    return row, col
#############################################################################


def calculate_col_total(startcell, endcell):
    colval = 0
    for x in startcell.expand('down'):
        colval = colval + x.value

    endcell.value = colval
    endcell.api.Font.Bold = True
    endcell.color = (255, 255, 0)
#############################################################################


def calculate_row_total(startcell, endcell):
    rowval = 0
    for x in startcell.expand('right'):
        rowval = rowval + x.value

    endcell.value = rowval
    endcell.api.Font.Bold = True
    endcell.color = (255, 255, 0)
#############################################################################


def border(sheet, a, b):
    """ This function adds border lines in the Excel sheet """

    for cell in sheet.range(a + ':' + b).current_region:
        for border_id in range(7, 12):
            cell.api.Borders(border_id).LineStyle = 1
            cell.api.Borders(border_id).Weight = 2
#############################################################################


if __name__ == "__main__":
    main(sys.argv[1])
    #main()
