#!/usr/bin/
#coding=gbk
'''
Created on 2013年10月10日

@summary: this file is used to read statics datas from mysql  

@author: huxiufeng
'''

import MySQLdb


class dbproc:
    def __init__(self, myhost, myusr, mypwd):
        self.myhost, self.myusr, self.mypwd = myhost, myusr, mypwd
        
    def conndb(self):
        try:
            self.cxn = MySQLdb.Connect(host = self.myhost, user = self.myusr, passwd=self.mypwd)
            if self.cxn is None:
                return False
            self.cur = self.cxn.cursor()
            if self.cur is None:
                return False
            self.excuteidx = 0
            return True
        except Exception, e:
            print "error", e
            return False
    
    def closedb(self):
        self.cur.close()
        self.cxn.close()
        
    def do_select(self, sql_string):
        if self.do_excute(sql_string):
            for row in self.cur.fetchall():
                yield row

    
    def do_excute(self, sql_string):
        try:
            self.excuteidx +=1
            self.cur.execute(sql_string)
            print self.excuteidx,': Excute "', sql_string, '" done.'
            return True
        except Exception, e:
            print self.excuteidx,': error when excute "', sql_string, '"\n' , e
            return False
        finally:
            pass


#----------------------It is a split line--------------------------------------

def main():
    sql_select_mtlog1 = r'select count(*), rec_crt_ts from msonldb.tbl_smsmt_log2 group by rec_crt_ts;'
    sql_excute_usedb = "use msonldb;"

    myhost = r'172.17.254.244'
    myusr = r'msonldb'
    mypwd = r'123'
    
    dbconn = dbproc(myhost, myusr, mypwd)
    if not dbconn.conndb():
        print "connect db error"
        return
    
    
    dbconn.do_excute(sql_excute_usedb)
    
    
    for row in dbconn.do_select(sql_select_mtlog1):
        print row
    
    dbconn.closedb()
    
#----------------------It is a split line--------------------------------------

if __name__ == "__main__":
    main()
    print "It's ok"