import xlrd
import openpyxl
from dateutil import parser
from datetime import datetime


def get_events(path, sheet_index=0, key="SF"):
    """Given a path of an excel file, and the key word to search for,
    return a list of row name (part title) and column name (date)"""

    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(sheet_index)
    
    wbk = openpyxl.load_workbook(path)
    wks = wbk.worksheets[0]

    events = []

    # Scan the excel file for all cells that contanin the key ("SF") and return them
    for i in range(1,sheet.nrows):
        desc = ""
        name = ""
        # print(sheet.cell_value(i,14),sheet.cell_value(i,14) != "TRUE")
        if sheet.cell_value(i,14) != "TRUE": #processed
            # print(sheet.cell_value(i,13))
            if sheet.cell_value(i,13) == 1: #interested
                for j in range(1, sheet.ncols - 2):
                    
                    if j == 1:
                        name = sheet.cell_value(i,j)
                    elif j!= 5:
                        desc+= sheet.cell_value(0,j) + ":" + sheet.cell_value(i,j)

                if i!= 0 and len(sheet.cell_value(i, 5))>0:
                    # print(sheet.cell_value(i, 5))
                    events.append([name,desc, str(parser.parse((sheet.cell_value(i, 5)+"11pm")).isoformat()),str(parser.parse((sheet.cell_value(i, 5)+"11:59pm")).isoformat())])
            wks.cell(row=i+1, column=14+1).value = "TRUE"
    wbk.save(path)
    wbk.close
    # print(events)
    # exit()
    return events


def main():
    loc = ("E:\Programs\Script\Create Events From Excel\Sound-Platform-Attendants Complete List_March 2019-1 (1).xlsx")

    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)

    for i in range(sheet.nrows):
        for j in range(sheet.ncols):
            if (sheet.cell_value(i, j) == 'Date'):
                # Storing the row index for the date (since it is not 0)
                # so the event date can be found on that corresponding row
                date_row = i
            if (sheet.cell_value(i, j) == 'SF'):
                # Alternate solution
                # myDate = datetime.strptime(sheet.cell_value(date_row, j)+"-2019", '%d-%b-%Y')

                # Since the date is in 'dd-MMM' format, we need to convert it to isoformat
                print(sheet.cell_value(i, 0), str(parser.parse(sheet.cell_value(date_row, j)).date()))


if __name__ == '__main__':
    main()

# Program extracting all columns
# name in Python
# import xlrd
#
# loc = ("E:\Programs\Script\Sound-Platform-Attendants Complete List_March 2019-1 (1).xlsx")
#
# wb = xlrd.open_workbook(loc)
# sheet = wb.sheet_by_index(0)
#
# # For row 0 and column 0
# sheet.cell_value(0, 0)
#
# for i in range(sheet.ncols):
#     print(sheet.cell_value(0, i))
