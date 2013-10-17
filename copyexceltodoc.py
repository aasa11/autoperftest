#!/usr/bin/
#coding=gbk
'''
Created on 2013.10.10

@summary: 

@author: huxiufeng
'''
import os
import xlsHelper
import docHelper



def copy_chart_from_xls_to_doc(xlspath, docpath, sheetname, chartidx, doctag=None):
    xls = xlsHelper.ExcelHelper(xlspath)
    xls.do_copychart(sheetname, chartidx)
    xls.close()
    
    doc = docHelper.docHelper(True)
    doc.open(docpath)
    doc.do_paste()
    doc.save()
    

def run_xls_marco(xlspath, sheetname, marconame, *paras):
    xls = xlsHelper.ExcelHelper(xlspath, True)
    xls.cal_Macro(sheetname, marconame, *paras)
    #xls.close()
    


#----------------------It is a split line--------------------------------------

def main():
    xlspath = os.path.abspath(r'res/224_linux-cck1_130930_0907.nmon.xlsx')
    docpath = os.path.abspath(r'res/doc_demo.doc')
    sheetname = r'DISKXFER'
    chartidx = 2
    copy_chart_from_xls_to_doc(xlspath, docpath, sheetname, chartidx)
    
    xlspath2 = os.path.abspath(r'res/nmon analyser v34a.xlsm')
    sheetname2 = r'Analyser'
    marconame = r'CommandButton1_Click'
    #marconame = r'Main'
#     xlspath2 = os.path.abspath(r'res/224_linux-cck1_130930_0907.nmon.xlsm')
#     sheetname2 = r'Sheet1'
#     marconame = r'marco12'
#     run_xls_marco(xlspath2, sheetname2, marconame)
    
#----------------------It is a split line--------------------------------------

if __name__ == "__main__":
    main()
    print "It's ok"