#!/usr/bin/
#coding=gbk
'''
Created on 2013年10月11日

@summary: 

@author: huxiufeng
'''
import os
import ConfigParser


class Mycnf:
    def __init__(self, filepath):
        self.cnf = ConfigParser.ConfigParser()
        abspath = os.path.abspath(filepath)
        if not os.path.exists(abspath):
            print "error path"
            return 
        self.cnf.read(abspath)
        ksv = self.cnf.sections()
        for v in ksv:
            print '\nsections: ',v
            print 'items: '
            for vv in self.cnf.items(v):
                print '\t ', vv
        #print ksv
        
        #section define
        self.sec_base = r'sec_base'
        self.sec_doc = r'sec_doc'
        self.sec_db = r'sec_db'
        self.sec_sql = r'sec_sql'
        self.sec_perfomance = r'sec_perfomance'
        self.sec_sshhost = r'sec_sshhost'
        
        #option define
        self.op_demo_path = r'demo_path'
        self.op_res_path = r'res_path'
        self.op_outer_path = r'outer_path'
        
        self.op_ext = r'ext'
        
        self.op_word_demo = r'word_demo'
        self.op_word_new = r'word_new'
        self.op_host = r'host'
        self.op_usr = r'usr'
        self.op_pwd = r'pwd'
        self.op_db = r'db'
        self.op_sql = r'sql'
        self.op_xls = r'xls'
        self.op_machineid = r'machineid'
        self.op_chart = r'chart'

        
    def get_base_path(self):
        return  self.cnf.get(self.sec_base, self.op_demo_path),  \
                self.cnf.get(self.sec_base, self.op_res_path)  
    
    def get_base_outer_path(self):    
        return  self.cnf.get(self.sec_base, self.op_outer_path)
    
    def get_db_info(self):
        return self.cnf.get(self.sec_db, self.op_host), \
                self.cnf.get(self.sec_db, self.op_usr), \
                self.cnf.get(self.sec_db, self.op_pwd), \
                self.cnf.get(self.sec_db, self.op_db),  \
                self.cnf.get(self.sec_db, self.op_sql)  
                
    def get_sql_name(self):
        return self.cnf.get(self.sec_sql,self.op_xls)
    
    def get_doc_info(self):
        return self.cnf.get(self.sec_doc,self.op_word_demo), \
                self.cnf.get(self.sec_doc,self.op_word_new)
                
    def get_sql_tag(self):
        return self.cnf.get(self.sec_sql,self.op_xls),\
                self.cnf.get(self.sec_sql,self.op_machineid),\
                self.cnf.get(self.sec_sql,self.op_chart)
                
    def get_perfomance_tag(self):
        return self.cnf.get(self.sec_perfomance,self.op_xls),\
                self.cnf.get(self.sec_perfomance,self.op_machineid),\
                self.cnf.get(self.sec_perfomance,self.op_chart)
            
    def get_sshfile_ext(self):
        return self.cnf.get(self.sec_sshhost, self.op_ext) 
    
    
    def get_sshfile_host(self):
        ksv = self.cnf.items(self.sec_sshhost)
        #print ksv
        
        paras = []
        i = 0
        for its in ksv:
            if its[0] == self.op_ext:
                continue
            outlist = str(its[1]).split(':')
            paras.append(outlist)
            i+=1    
        return paras




#----------------------It is a split line--------------------------------------

def main():
    filepath  = r'cnf/cnf_report.ini'
    cnf = Mycnf(filepath)
    
#----------------------It is a split line--------------------------------------

if __name__ == "__main__":
    main()
    print "It's ok"