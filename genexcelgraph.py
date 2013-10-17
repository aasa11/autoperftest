#!/usr/bin/
#coding=gbk
'''
Created on 2013年10月10日

@summary: 

@author: huxiufeng
'''
import os
from win32com.client import  Dispatch
import xlsHelper

def excelChart(xlpath):
    xl = Dispatch("Excel.Application")
    
    print "xl = ", xl
    
    wb = xl.Workbooks.open(xlpath)
    
    xl.Visible = 1
    ws  = wb.Worksheets(1)
    ws.Range('$A1:$D1').Value = ['NAME', 'PLACE', 'RANK', 'PRICE'];
    ws.Range('$A2:$D2').Value = ['Foo', 'Fooland', 1, 100];
    ws.Range('$A3:$D3').Value = ['Bar', 'Barland', 2, 75];
    ws.Range('$A4:$D4').Value = ['Stuff', 'Stuffland', 3, 50];
    wb.Save();
    wb.Charts.Add();
    #wc1 = wb.Charts(1);
    
    
def excelChart_uselib(xlpath):
    xls = xlsHelper.ExcelHelper(xlpath)
    datalst = [["col1", "col2", "col3"], [1,2,3], [3,4,5],[5,6,7],[9,8,7],[6,5,4]]
    xls.AddCharts(datalst)

#----------------------It is a split line--------------------------------------

def main():
    xlpath = r'res/chart_demo.xls'
    abspath = os.path.abspath(xlpath)
    excelChart_uselib(abspath)
    
#----------------------It is a split line--------------------------------------

if __name__ == "__main__":
    main()
    print "It's ok"