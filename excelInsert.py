import xlrd
from xlutils.copy import copy
from xlwt import Workbook
import xlwings as xw

from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showinfo

from icon import Icon

import base64
import os


'''
def xls2xlsx(src_file_path, dst_file_path):
    book_xls = xlrd.open_workbook(src_file_path)
    book_xlsx = xl.workbook.Workbook()
    sheet_names = book_xls.sheet_names()
    for sheet_index, sheet_name in enumerate(sheet_names):
        sheet_xls = book_xls.sheet_by_name(sheet_name)
        if sheet_index == 0:
            sheet_xlsx = book_xlsx.active
            sheet_xlsx.title = sheet_name
        else:
            sheet_xlsx = book_xlsx.create_sheet(title = sheet_name)

        for row in range(0, sheet_xls.nrows):
            for col in range(0, sheet_xls.ncols):
                sheet_xlsx.cell(row = row+1, column= col+1).value = sheet_xls.cell_value(row, col)
    book_xlsx.save(dst_file_path)
    return book_xlsx
'''


def insert_excel(inserted_num_str, existed_excel_name, inserted_excel_name):

    inserted_num = int(inserted_num_str)
    wb1 = xw.Book(existed_excel_name)
    wb2 = xw.Book(inserted_excel_name)

    sheet_num = len(wb2.sheets)

    true_inserted_num = 0

    # insert w2 sheets
    for i in range(0, sheet_num):
        temp = wb2.sheets(i+1)
        sheet_old_name = temp.name
        if ("卷内单" in sheet_old_name) or ("卷内多" in sheet_old_name):
            continue
        if "—" in sheet_old_name:
            sheet_new_name = str(inserted_num+i) + "—" + sheet_old_name.split("—")[1]
            temp.name = sheet_new_name
            temp.range('B2').value = "档案号:" + str(inserted_num+i) + "——" + sheet_old_name.split("—")[1]
        temp.api.Copy(Before=wb1.sheets(inserted_num+i).api)
        true_inserted_num += 1
        wb1.save()

    orig_sheet_num = len(wb1.sheets)

    # change w1 sheets names if needed
    for i in range(inserted_num+true_inserted_num-1, orig_sheet_num):
        temp = wb1.sheets(i+1)
        sheet_old_name = temp.name
        if ("卷内单" in sheet_old_name) or ("卷内多" in sheet_old_name):
            continue
        if "—" in sheet_old_name:
            sheet_new_name = str(i+1) + "—" + sheet_old_name.split("—")[1]
            temp.name = sheet_new_name
            temp.range('B2').value = "档案号:" + str(i+1) + "——" + sheet_old_name.split("—")[1]

        wb1.save()

    wb1.app.quit()

'''
    #获取inserted_excel_name中的所有sheet
    inserted_book = xlrd.open_workbook(inserted_excel_name, formatting_info=True)
    for sheet in inserted_book.worksheets:
        new_sheet = existed_book.create_sheet(sheet.title)
        for row in sheet:
            for cell in row:
                new_sheet[cell.coordinate].value = cell.value

    existed_book.save(existed_excel_name)

'''
def GUIfileopen(value):
    file_excel = askopenfilename()
    if file_excel:
        value.set(file_excel)


if __name__ == '__main__':

    #ICON
    with open('tmp.ico', 'wb') as tmp:
        tmp.write(base64.b64decode(Icon().img))

    #GUI
    frameT = Tk()
    frameT.geometry('500x180+400+200')
    frameT.title("你家小可爱的小程序啦啦啦啦~~~~")
    frameT.wm_iconbitmap('tmp.ico')
    os.remove('tmp.ico')

    #inserted_num
    frame_num = Frame(frameT)
    frame_num.pack(padx=10, pady=10)
    #original file
    frame_orig = Frame(frameT)
    frame_orig.pack(padx=10, pady=10)
    #inserted file
    frame_insert = Frame(frameT)
    frame_insert.pack(padx=10, pady=10)
    #exe
    frame_exe = Frame(frameT)
    frame_exe.pack(padx=10, pady=10)


  # global var
    inserted_num = IntVar()
    existed_excel_name = StringVar()
    inserted_excel_name = StringVar()

    label1 = Label(frame_num, width=40, text="要插在第几个表单前面呐？", font=("宋体", 9))
    label1.pack(fill=X, side=LEFT)

    entry1 = Entry(frame_num, width=10, textvariable=inserted_num)
    entry1.pack(fill=X, side=RIGHT)
    entry2 = Entry(frame_orig, width=40, textvariable=existed_excel_name)
    entry2.pack(fill=X, side=LEFT)
    entry3 = Entry(frame_insert, width=40, textvariable=inserted_excel_name)
    entry3.pack(fill=X, side=LEFT)

    btn_orig = Button(frame_orig, width=40, text="选择已归档文件名", font=("宋体", 9),
                      command=lambda: GUIfileopen(existed_excel_name)).pack(fill=X, padx=10)
    btn_insert = Button(frame_insert, width=40, text="选择要插入的文件名", font=("宋体", 9),
                        command=lambda: GUIfileopen(inserted_excel_name)).pack(fill=X, padx=10)
    btn_exe = Button(frame_exe, width=20, text="开始工作啦！", font=("宋体", 9),
                     command=lambda: insert_excel(entry1.get(), entry2.get(), entry3.get())).pack(fill=X, padx=10)
    frameT.mainloop()

    #input
    #inserted_num = 3 #插在第1个文件后面

    #existed_excel_name = 'G:\档案工作\\2018归档记录\\2018移交目录\\2018.11.16\\青岛地铁1号线卷内目录.xls'
    #inserted_excel_name = 'G:\档案工作\\2018归档记录\\2018移交目录\\2018.11.16\\黄石海洲卷内目录.xls'

    #将inserted_excel_name插入existed_excel_name的第inserted_num sheet前面
    #insert_excel(inserted_num, existed_excel_name, inserted_excel_name)