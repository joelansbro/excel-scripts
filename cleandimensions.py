# iterate dimensions columns and if the weights are decimalised rounds them up
# sorts them in following order =MAX =MIN =MEDIAN (length height width)
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook

# gets col header based on parameter
def findCol(col_name, ws):
    for row in range(1, ws.max_column):
        if(ws.cell(1,row).value == col_name):
            return ws.cell(1,row)

# sorts the length height and width into max, min, median
def sortDims(length, height, width):
    if(length < width):
        length, width = width, length
    if(length < height):
        length, height = height, length
    if(width < height):
        width, height = height, width
    return(length, height, width)    



# this dim sort is being fed coordinates eg ['C1'], ['D1'], ['E1']
# take the column and iterate down 
# for row in C[i] -> store in list
#    then 
def dimSort(length, height, width, ws):
    print(length, height, width)
    for row in range(2,ws.max_row):
        print(str(length) + str(row))
        print(str(height) + str(row))
        print(str(width) + str(row))  
        # we have accessed the needed cells, now you must sort them into a list
        # once in the list, run sortdims
        # once returned, pass the sorted dimensions to the proper cells that they came in with  



if __name__ in '__main__':
    wb = openpyxl.load_workbook(input("absolute path of file") + ".xlsx", read_only=False)
    ws = wb.active

    print("This programme will sort the Length Height and Width columns to be workable within CCP\n")

    VAR_length = findCol(input("Type the name of the Length column \n"), ws)
    VAR_height = findCol(input("Type the name of the Height column \n"), ws)
    VAR_width = findCol(input("Type the name of the Width column \n"), ws)
    print("Length = " + VAR_length.value)
    print("Height = " + VAR_height.value)
    print("Width = " + VAR_width.value)
    dimSort(VAR_length.column_letter, VAR_height.column_letter, VAR_width.column_letter, ws)
    # sortDims(1,2,3)