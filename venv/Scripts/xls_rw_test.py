import xlrd
import xlwt

filename = r'D:\JetBrains\xls_RW\广东省资源清单\1.xlsx'

book1 = xlrd.open_workbook(filename)
book2 = xlwt.Workbook(encoding='ascii')
sheet2 = book2.add_sheet("copy from 1.xls")

sheetlist = book1.sheet_names()
# print(sheetlist)

book_info ={}
for _ in (sheetlist):
    sheet = book1.sheet_by_name(_)
    print("Add %s Row %d --- Column %d into Book_info"%( _, sheet.nrows, sheet.ncols))
    book_info[_]=[sheet.nrows,sheet.ncols]


#for _ in book_info.keys():
#    print ("KEY %s Row %d --- Column %d"%(_, book_info[_][0], book_info[_][1]))

sheet0 = book1.sheet_by_index(0)

cur_row = 0
row=[0,0,0,0]
rw_row =0
while (cur_row<sheet0.nrows):
    cur_col =0
    new=0
    while(cur_col < sheet0.ncols and cur_col < 4):
        if (row[cur_col] != sheet0.cell_value(cur_row, cur_col)):
            row[cur_col] = sheet0.cell_value(cur_row, cur_col)
            new=1
        cur_col += 1
    if (new == 1):
        i=0
        for _ in (row):
            sheet2.write(rw_row, i, _)
            i+=1
        rw_row +=1
    cur_row+=1

book2.save('D:\\JetBrains\\xls_RW\广东省资源清单\\2.xls')
    #print (sheet.row(cur_row))