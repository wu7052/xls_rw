from xls_rw_test import xls_rw, qty_split
#from wxpy import *

#print ("__name__ %s" % __name__)
if (__name__ == "__main__"):
    print ("start to QTY_SPLIT")
    main_file = r'D:\02龙华\LH智慧校园\按学校拆分\0917.xls'
    src_file = r'D:\02龙华\LH智慧校园\按学校拆分\龙华中心小学.xlsx'

    #if (qty_split(main_file,src_file) == -1):
    if (qty_split(main_file, src_file) == -1):
        print ("文件操作有问题，退出了\n")
        exit(0)

    # xls_rw()
    # bot = Bot()
    # my_friend = bot.friends().search('Melissa')[0]
    # my_friend.send('睡着了？')
