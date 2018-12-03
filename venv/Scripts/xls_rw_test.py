import xlrd
import xlwt
from xlutils.copy import copy
import os
import re


def test_file(filename=None):
    if filename is None:
        print("Filename is needed .")
        return False
    if not os.path.isfile(filename):
        print("%s is NOT exist" % filename)
        return False
    if not os.access(filename, os.R_OK):
        print("%s is NOT accessible to read" % filename)
        return False
    if not os.access(filename, os.W_OK):
        print("%s is NOT accessible to write" % filename)
        return False
    return True


def qty_split(main_file=None, src_file=None):
    if not test_file(main_file):
        return -1
    if not test_file(src_file):
        return -1
    sheet_list = ['文化视窗', '互动课堂', '录播', '物联', '安全校园']

    src_book = xlrd.open_workbook(src_file)
    main_book = xlrd.open_workbook(main_file)

    main_book_wt = copy(main_book)  # 复制页面
    #print(src_book.sheet_names())
    #print(main_book.sheet_names())

    src_sheet_list = src_book.sheet_names()
    main_sheet_list = main_book.sheet_names()

    sheet_map = {}
    i = 0
    while len(sheet_list) > 0:
        sheet_name = r'\w*' + sheet_list[0] + r'\w*'
        for _ in main_sheet_list:
            if re.match(sheet_name, _):
                main_sheet_name = _
                main_sheet_list.remove(_)

        for _ in src_sheet_list:
            if re.match(sheet_name, _):
                src_sheet_name = _
                src_sheet_list.remove(_)

        sheet_map[i] = [main_sheet_name, src_sheet_name]

        sheet_list.pop(0)
        i += 1

    print(sheet_map)

    """
    for _ in main_book.sheet_names():
        i =0
        while (i<5):
            sheet_name = r'\w*'+sheet_list[i]+r'\w*'
            if re.match(sheet_name, _):
                print ("[%s] is found in sheet_list[%d]"%(_, i))
                # print ("[%s] start to be processed... " % _)
                # sheet_process(_, main_book_wt, src_book)
            # print ("%d, %s " % (i,  sheet_name))
            i+=1
    """


def sheet_process(sheet_map=None, main_book_handle=None, src_book_handle=None):
    if sheet_map is None:
        print("[sheet_proces] need a sheet map \n")
        return -1

    if main_book_handle is None:
        print("[sheet_proces] need a main sheet handle \n")
        return -1

    if src_book_handle is None:
        print("[sheet_proces] need a source sheet handle \n")
        return -1



# main_book_sheet =
"""
    for i in range(96):
        ws.write(1, 5 + i, vector[i])
    # ----- 按(row, col, str)写入需要写的内容 -------
    main_book_wt.save(main_file) 		# 保存文件
"""


def xls_rw():
    filename = r'D:\JetBrains\xls_RW\广东省资源清单\1.xlsx'

    book1 = xlrd.open_workbook(filename)
    book2 = xlwt.Workbook(encoding='ascii')
    sheet2 = book2.add_sheet("copy from 1.xls")

    sheetlist = book1.sheet_names()
    # print(sheetlist)

    book_info = {}
    for _ in sheetlist:
        sheet = book1.sheet_by_name(_)
        print("Add %s Row %d --- Column %d into Book_info" % (_, sheet.nrows, sheet.ncols))
        book_info[_] = [sheet.nrows, sheet.ncols]

    # for _ in book_info.keys():
    #    print ("KEY %s Row %d --- Column %d"%(_, book_info[_][0], book_info[_][1]))

    sheet0 = book1.sheet_by_index(0)

    cur_row = 0
    row = [0, 0, 0, 0]
    rw_row = 0
    while cur_row < sheet0.nrows:
        cur_col = 0
        new = 0
        while cur_col < sheet0.ncols and cur_col < 4:
            if row[cur_col] != sheet0.cell_value(cur_row, cur_col):
                row[cur_col] = sheet0.cell_value(cur_row, cur_col)
                new = 1
            cur_col += 1
        if (new == 1):
            i = 0
            for _ in row:
                sheet2.write(rw_row, i, _)
                i += 1
            rw_row += 1
        cur_row += 1

    book2.save('D:\\JetBrains\\xls_RW\广东省资源清单\\3.xls')
    # print (sheet.row(cur_row))
