#!/usr/bin/
#coding=gbk
'''
Created on 2013年10月11日

@summary: 

@author: huxiufeng
'''
import os
import shutil
import time
import xlsHelper
import docHelper
import cnf_reader
import getmysqldata

#gen tps chart
def gen_tps_chart(cnf):
    myhost, myusr, mypwd, mydb, mysql = cnf.get_db_info()
    _, xlspath = cnf.get_base_path()
    apdx = cnf.get_sql_name()
    xlspath   = os.path.join( os.path.abspath(xlspath), apdx+'.xlsx')
    
    print xlspath
    
    xls = xlsHelper.ExcelHelper(xlspath)
    datalist = []
    
    dbconn = getmysqldata.dbproc(myhost, myusr, mypwd)
    if not dbconn.conndb():
        print "connect db error"
        return False
    for row in dbconn.do_select(mysql):
        print row
        datalist.append(list(row))
    
    dbconn.closedb()

    if len(datalist)> 0:
        xls.AddCharts(datalist)
        xls.do_copyhere()
    xls.save()
    xls.close()
    

def copy_doc(cnf):
    demopath, docpath = cnf.get_base_path()
    demoname, docname = cnf.get_doc_info()
    demoname = os.path.join(os.path.abspath(demopath),demoname)
    docname = os.path.join(os.path.abspath(docpath),docname)
    print demoname
    print docname
    
    if os.path.exists(demoname):
        print "copy file"
        shutil.copy(demoname,  docname)
    return docname

def get_tag(machineid, chartinfo):
    tag2 = chartinfo.split(',')[0].strip()
    tag = r'<'+machineid+'_'+tag2+r'>'
    return tag

def get_tag_lst(machineid, lst):
    tag = r'<'+machineid+'_'+lst[0]+r'>'
    return tag

def copy_sql_chart_to_doc(cnf, docname, hascopydata = True):
    doc = docHelper.docHelper(True)
    doc.open(docname)
    
    nameinfo, machineid, chart_info = cnf.get_sql_tag()
    sql_tag = get_tag(machineid, chart_info)
    print sql_tag
    
    if hascopydata:
        doc.do_paste(sql_tag)
    else:
        print "do not realizeed , todo later..."
        #todo
        pass
    doc.save()
    doc.close()
    
def copy_performance_charts_to_doc(cnf, docname):
    demopath, docpath = cnf.get_base_path()
    nameinfo, machineid, chart_info = cnf.get_perfomance_tag()
    
    namelist = list(nameinfo.split(','))
    machineidlist = list(machineid.split(','))
    chartlist = []
    l1 = list(chart_info.split(';'))
    for v in l1:
        if v == '':
            break
        l2 = list(v.split(','))
        chartlist.append(l2)
        
    name_machine = zip(namelist, machineidlist)
    print name_machine,chartlist
        
    docpath = os.path.abspath(docpath)
    
    doc = docHelper.docHelper(False)
    doc.open(docname)
    time.sleep(0.5)
    
    for v in  name_machine:
        namepart = v[0]
        machineid = v[1]
        filename = ''
        for _,_,files in os.walk(docpath):
            for f in files:
                if v[0] in f and (f.endswith('.xls') or f.endswith('.xlsx')):
                    print "find ", namepart, ", name is ", f
                    filename = os.path.join(docpath,f)
                    break
            break
        if filename == '':
            continue
        xls = xlsHelper.ExcelHelper(filename)
        for chartv in chartlist:
            print "copy chart file:", filename, " , sheet: ", chartv[1], " , chartid : ", chartv[2]
            if xls.do_copychart(str(chartv[1]).strip(), int(chartv[2])):
                doc_tag = get_tag_lst(machineid, chartv)
                print "doc_tag " , doc_tag
                isok = False
                for _ in range(3):
                    time.sleep(2)
                    if doc.do_paste(doc_tag):
                        isok = True
                        break      
                if not isok:
                    print "paste err for retry three times...."           
        xls.close()
        
        
    doc.save()
    time.sleep(1)
    doc.visible(True)
    #doc.close()
    

def do_stat(cnffile):
    #read conf file
    cnf = cnf_reader.Mycnf(cnffile)
    
    #generate db tps chart in excel
    gen_tps_chart(cnf)
   
    #copy new doc from demo     
    docname = copy_doc(cnf)
    #copy tps charts to doc
    copy_sql_chart_to_doc(cnf, docname)
    
    #copy performance charts to doc
    copy_performance_charts_to_doc(cnf, docname)
    



#----------------------It is a split line--------------------------------------

def main():
    cnffile = r'cnf/cnf_report.ini'
    do_stat(cnffile)
    
    
    
#----------------------It is a split line--------------------------------------

if __name__ == "__main__":
    main()
    print "It's ok"