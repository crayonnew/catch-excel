import xlrd
import tkinter
import re
from xlutils.copy import copy


def get_path():
    if text.get () != '':
        words = text.get()
        workbook = xlrd.open_workbook (words)
        table = workbook.sheets ()[0]
        n = 1
        while True:
            try:
                col_values = table.cell (n, 33).value
                if "for" in col_values or "for lining" in col_values:
                    row_values = re.findall ("\sin\s(\w+\s?\w*)\s?f+", col_values)
                else:
                    row_values = re.findall ("\sin\s(\w+\s?\w*)\s?", col_values)
            except IndexError:
                break
            n += 1
            for value in row_values:
                lists.append (value)
            old_book = xlrd.open_workbook (words)
            new_book = copy (old_book)
            dec_table = new_book.get_sheet (0)
            num = 1
            for list in lists:
                dec_table.write (num, 40, list)
                num += 1
            try:
                new_book.save ("new" + words)
            except PermissionError:
                hint = tkinter.Label (root, text='错误：文件已经打开', font=("宋体", 10, "normal"))
                hint.place (x=47, y=80, width=150, height=12)
            else:
                hint = tkinter.Label (root, text='抓取完成', font=("宋体", 10, "normal"))
                hint.place (x=80, y=60, width=80, height=12)


lists = list()
word_lists = list()
root = tkinter.Tk ()
root.title ("自动抓取excel")
root.geometry ('300x120')
root.resizable (False, False)
root.flag = True

text = tkinter.Entry (root)
text.place(x=60, y=30, width=130, height=20)
comand = tkinter.Button (root, text="添加", command=get_path)
comand.place (x=200, y=30, width=80, height=20)

hint = tkinter.Label (root, text='输入文件名', font=("宋体", 10, "normal"))
hint.place (x=80, y=10, width=80, height=12)

hint = tkinter.Label (root, text='文件放在同级目录下', font=("宋体", 10, "normal"))
hint.place (x=60, y=100, width=120, height=12)

root.mainloop ()

