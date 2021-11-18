import datetime,sys,os.path,openpyxl,os,keyboard,time
from termcolor import colored, cprint
from datetime import datetime

def openfile(header):
    file_name = datetime.datetime.now().strftime('%d_%m_%Y')+'.xlsx'
    print(file_name)
    PATH = './log/'+file_name
    if os.path.isfile(PATH):
    # print('da ton tai')
        wb = openpyxl.load_workbook(filename=PATH)
        # ws = wb.active
    else:
    # print('chau ton tai')
        wb = openpyxl.Workbook()
        # ws = wb.active
        # ws.title = "NHAT KY NGAY"
        # ws.append(header[1:])
    return ws

# openfile()