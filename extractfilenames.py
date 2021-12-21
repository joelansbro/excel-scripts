# this script searches a directory and returns the names of all the files within a CSV

import csv
import os

filePath = input("Enter the filename below\n")

f = csv.writer(open("Filelist.csv", "W"))

directories = os.walk(filepath, topdown=True)
for root, dirs, files in directories:
    for file in files:
        f.writerow([file])
        itemID = file[0:5]
        newFile = itemID