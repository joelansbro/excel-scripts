import openpyxl
import pandas as pb

df = pb.read_excel(input("Enter the absolute path of the file: \n"))

def dimensions():
    print("Enter the dimensions columns")
    zero = df[['VAR_Weight', 'VAR_Length', 'VAR_Height','VAR_Width']].eq(0).sum()
    print(zero)

def emptyVals():
    print("Enter the value columns")
    category = df[['RNG_SKU','RNG_Name','OPT_colour']].isnull().sum()
    print(category)



if __name__ in '__main__':
    menu = True
    while(menu):
        status = input("Enter dimensions, values, or exit")
        match status:
            case "dimensions":
                dimensions()
            case "values":
                emptyVals()
            case "exit":
                print("Exiting programme")
                menu = False
            case _:
                print("Invalid value, expecting Dimensions, Values, or Exit")
                