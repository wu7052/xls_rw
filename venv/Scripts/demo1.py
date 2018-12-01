from xls_rw_test import xls_rw
from wxpy import *

#print ("__name__ %s" % __name__)
if (__name__ == "__main__"):
    print ("start to xls_rw")
    # xls_rw()
    bot = Bot()
    my_friend = bot.friends().search('Melissa')[0]
    my_friend.send('睡着了？')
