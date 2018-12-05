from typing import Dict, List, Any, Union

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
    main_book = xlrd.open_workbook(main_file, formatting_info=True)

    main_book_wt = copy(main_book)  # 复制页面
    # print(src_book.sheet_names())
    # print(main_book.sheet_names())

    src_sheet_list = src_book.sheet_names()
    main_sheet_list = main_book.sheet_names()

    sheet_map: Dict[ Union[int, Any], List[Any] ] = {}
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

        """
        sheet_map 字典格式
        {0:['目标文件 sheet名称','学校文件 sheet名称']; 
         1:['目标文件 sheet名称','学校文件 sheet名称'];
         2:['目标文件 sheet名称','学校文件 sheet名称'];
         3:['目标文件 sheet名称','学校文件 sheet名称'];
         4:['目标文件 sheet名称','学校文件 sheet名称'];}
        """
        sheet_map[i] = [main_sheet_name, src_sheet_name]

        sheet_list.pop(0)
        i += 1

    print(sheet_map)
    sheet_process(sheet_map, main_book,
                  main_book_wt, src_book, main_file)


def sheet_process(sheet_map=None, main_book_rd=None,
                  main_book_wt=None, src_book_rd=None, main_file=None):

    if sheet_map is None:
        print("[sheet_proces] need a sheet map \n")
        return -1

    if main_book_rd is None:
        print("[sheet_proces] need a main sheet handle \n")
        return -1

    if main_book_wt is None:
        print("[sheet_proces] need a main sheet handle \n")
        return -1

    if src_book_rd is None:
        print("[sheet_proces] need a source sheet handle \n")
        return -1

    src_map = {}

    # 目标表单 写入操作 句柄
    m_sheet_wt = main_book_wt.get_sheet(sheet_map[0][0])

    # 目标表单 读取操作 句柄
    m_sheet_rd = main_book_rd.sheet_by_name(sheet_map[0][0])
    # print("sheet name:{name}--- ncolumn:{col} --- nrow:{row}".
    #     format(name=src_book_rd.sheet_names()[0], row=m_sheet_rd.nrows, col=m_sheet_rd.ncols))

    # 在目标表单 写入学校名，并记录 列号
    m_sheet_wt.write(0, m_sheet_rd.ncols, src_book_rd.sheet_names()[0])
    sheet_map[0].append(m_sheet_rd.ncols)
    print (sheet_map)

    # 源表单 读取操作句柄
    src_sheet_rd = src_book_rd.sheet_by_name(sheet_map[0][1])
    src_map = src_map_gather(src_sheet_rd)

    main_book_wt.save(main_file)


def src_map_gather(src_sheet_rd=None):
    if src_sheet_rd is None:
        print ("[src_map_gather] source sheet read is None , return]")
        raise excepiton
    cur_row , content_flag= 0 ,0
    target_column_index = {'序号':0, '设备':0, '数量':0}
    target_map={}
    target_map_index = 0
    while cur_row < src_sheet_rd.nrows:
        row_content = src_sheet_rd.row_values(cur_row)
        index = 0

        # 确定目标列的序号
        if content_flag == 0:
            while index < len(row_content):
                if re.match((r'.*序号.*'),row_content[index]):
                    target_column_index['序号'] = index
                    content_flag = 1
                elif re.match((r'.*设备.*'), row_content[index]):
                    target_column_index['设备'] = index
                    content_flag = 1
                elif re.match((r'.*名称.*'), row_content[index]):
                    target_column_index['设备'] = index
                    content_flag = 1
                elif re.match((r'.*数量.*'), row_content[index]):
                    target_column_index['数量'] = index
                    content_flag = 1
                index += 1
            """ 测试 列表的index 方法，返回index
            try:
                target_column_index['序号'] = row_content.index(r'序号')
            except Exception as e:
                pass
            else:
                print("found !")
                content_flag =1
            """
        # 已确定目标列，开始逐行读取目标列 写入 map 字典
        else :
            #print ("[Target Column] {}".format(target_column_index))
            target_map[target_map_index]=[row_content[target_column_index['序号']],row_content[target_column_index['设备']],
                                                                  row_content[target_column_index['数量']]]
            target_map_index += 1
        cur_row += 1

    # print (target_map)
    return target_map
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
        if new == 1:
            i = 0
            for _ in row:
                sheet2.write(rw_row, i, _)
                i += 1
            rw_row += 1
        cur_row += 1

    book2.save('D:\\JetBrains\\xls_RW\广东省资源清单\\3.xls')
    # print (sheet.row(cur_row))
