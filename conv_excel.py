#!/usr/bin/python3
# AKIMOTO
# 2021-08-06

import re, os, sys, openpyxl

if len(sys.argv) == 1 :
    print("USAGE")
    print("$ %s [file]" % sys.argv[0].split("/")[-1])
    quit()

for i in range(1, len(sys.argv)):
    
    input_file = sys.argv[i]
    
    file_open = open(input_file, "r").readlines()

    # excel
    excel_open       = openpyxl.Workbook()
    excel_sheet      = excel_open.active
    excel_sheet.title = "sheet1"
    excel_file_name  = "%s.xlsx" % input_file.split(".")[0]

    row_num = 0
    for file_line in file_open :

        sharp = "#" in file_line

        #if sharp == False : #True :
        #    continue
        #else :
        if True :
            row_num += 1
            column_num = len(file_line.split()) + 1
            data       = file_line.split()

            for columns in range(1,column_num):
                
                try :
                    insert_data = float(data[columns-1])
                except ValueError :
                    insert_data = str(data[columns-1])

                excel_sheet.cell(row=row_num, column=columns).value = insert_data

    excel_open.save(excel_file_name)
    print("output; %s" % excel_file_name)
