from pathlib import Path as p
import openpyxl
from openpyxl import utils

loc_excel_docs = p.cwd().joinpath('ExcelDocs')  # Create a path to directory with documents

wb = openpyxl.load_workbook(p.joinpath(loc_excel_docs, 'responses.xlsx'))  # Open excel

rspns_sheet = wb['Form Responses 1']
rspns = []  # list with spread sheet response each row in a dictionary

dic_temp = {}

max_cols = rspns_sheet.max_column  # get maximum number of columns
max_rows = rspns_sheet.max_row  # get maximum number of rows

for row in range(2, max_rows):
    for i in range(2, max_cols):
        var = utils.get_column_letter(i)  # convert column num to letter

        # info to get the keys for temp dictionary
        title = (rspns_sheet[(var + '1')].value[2:].strip(
            '.').strip())  # title for column( removes '.' white space and first 2 letters
        # dbs_key = None
        if '.' in title[:2]:
            """"
            if '.' in first two items of returned string return variable without the the '.' and use strip()
            """
            dbs_key = title[2:].strip()
        else:
            dbs_key = title

        # info to get values
        value = rspns_sheet[(var + str(row))].value

        dic_temp[dbs_key] = value  # add title & value to dictionary

    rspns.append(dic_temp.copy())
    dic_temp.clear()  # clear dictionary to prevent duplicate data

with open("Responses.txt",'a') as r:
    for key, value in rspns[1].items():
        r.write(key)
        r.write('\n')
#for x in rspns:
#r.write((str(x['Name of Church within the Parish where you attend Mass frequently'].strip() )))
#r.write(('\n'))
#print(x['Name of Church within the Parish where you attend Mass frequently'].strip())

#print("\n"*9)

#for    key , value in rspns[1].items():
#print(key,"--------> ", value)