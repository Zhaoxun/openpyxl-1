# this is to improve the ease to write values by openpyxl
# simply copy and paste codes below or from openpyxl_writers import *
# Author = 阎兆珣 (Zhaoxun Yan)

# row and column index or start follow excel convention starting from 1 (not 0)

# I ) write a single line with data in the form of python list

# Write one row in a particular sheet starting from particular column(default 1)
def write_row( datalist, sheetobj, rowindex, colstart=1):
    cur_col = colstart
    for data in datalist:
        sheetobj.cell(row=rowindex, column=cur_col).value = data
        cur_col += 1

# Write one column in a particular sheet starting from particular row(default 1)
def write_col( datalist, sheetobj, colindex, rowstart=1):
    cur_row = rowstart
    for data in datalist:
        sheetobj.cell(row=cur_row, column=colindex).value = data
        cur_row += 1

# II) Write multiple lines with data in the form of list on list
# list on list example: [range(1,5), range(6,8)] not necessarily in rectangle

# Flush rows by data as list on list from the upper-left cell
def flush_rows( listonlist, sheetobj, rowstart=1, colstart=1):
    cur_row = rowstart
    for line in listonlist:
        write_row(line, sheetobj, cur_row, colstart)
        cur_row += 1

# Flush columns by data as list on list from the upper-left cell
def flush_cols( listonlist, sheetobj, rowstart=1, colstart=1):
    cur_col = colstart
    for line in listonlist:
        write_col(line, sheetobj, cur_col, rowstart)
        cur_col += 1

# reading data can be abtained by library "xlrd" hence omitted here
# note xlrd follows python convention starting row and column from index 0
