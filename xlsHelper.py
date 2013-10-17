#!usr/bin/python
#coding=GBK
'''
Created on 2012-3-14
@author: huxiufeng
xls读取写入的包装类，用于读写xls
'''
import win32com.client
import os

class ExcelHelper:
    def __init__(self, filename=None, Visible = False):
        self.xlApp = win32com.client.Dispatch('Excel.Application')  
        self.xlApp.Visible = Visible
        if filename:
            self.filename=filename
            if os.path.exists(self.filename):
                self.xlBook=self.xlApp.Workbooks.Open(filename)
            else:
                self.xlBook= self.xlApp.Workbooks.Add()
        else:
            self.xlBook= self.xlApp.Workbooks.Add()
            self.filename='Untitle'
            
    def printerr(self, e):
        for v in e:
            #print type (v)
            if type(v) == type([]) or type(v) == type((1,2)):
                for vv in v:
                    print "\t\t", vv
            else:
                print "\t", v 
    
    def save(self, newfilename=None):  
        if newfilename:     
            self.filename = newfilename  
        self.xlBook.SaveAs(self.filename)    
        
    def saveandclose(self):
        self.xlBook.Close(SaveChanges=1)
        del self.xlApp
        
    def close(self):  
        self.xlBook.Close(SaveChanges=0)  
        del self.xlApp  
    
    def copySheet(self, before):  
        '''copy sheet'''  
        shts = self.xlBook.Worksheets  
        shts(1).Copy(None,shts(1))
    
    def newSheet(self,newSheetName):
        sheet=self.xlBook.Worksheets.Add()
        sheet.Name=newSheetName
        sheet.Activate()
        
    def getSheetCount(self):
        return self.xlBook.Worksheets.Count
    
    def getSheet(self, Index):
        if Index >= self.getSheetCount()+1 or Index <= 0:
            Index = 1
        sht = self.xlBook.Worksheets(Index)
        sht.Activate()
        return self.xlApp.ActiveSheet
    
    def activateSheet(self,sheetName):
        self.xlBook.Worksheets(sheetName).Activate()
        
    def activeSheet(self):
        return self.xlApp.ActiveSheet;    
    
    def getCell(self, row, col,sheet=None):  
        '''Get value of one cell''' 
        if sheet:
            sht = self.xlBook.Worksheets(sheet)  
        else:
            sht=self.xlApp.ActiveSheet    
        return sht.Cells(row, col).Value  
    
    def setCell(self, row, col, value,sheet=None):  
        '''set value of one cell'''
        if sheet:
            sht = self.xlBook.Worksheets(sheet)  
        else:
            sht=self.xlApp.ActiveSheet    
         
        sht.Cells(row, col).Value = value  
        
    def getRange(self, row1, col1, row2, col2,sheet=None):  
        '''return a 2d array (i.e. tuple of tuples)''' 
        if sheet:
            sht = self.xlBook.Worksheets(sheet)  
        else:
            sht=self.xlApp.ActiveSheet    
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value 
     
    def mergeCell(self, row1, col1, row2, col2,sheet=None): 
        if sheet:
            sht = self.xlBook.Worksheets(sheet)  
        else:
            sht=self.xlApp.ActiveSheet   
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Merge() 
    def rowsCount(self):
        '''return used rows count'''
        sht=self.activeSheet()
        return  sht.UsedRange.Rows.Count
    
    def GetUsedGange(self):
        sht = self.xlApp.ActiveSheet
        return sht.UsedRange
    
    def GetUsedColsNum(self):
        rng = self.GetUsedGange()
        return rng.Columns.Count    
    
    def GetUsedRowsNum(self):
        rng = self.GetUsedGange()
        return rng.Rows.Count 
    
    def GetSheetName(self):
        return self.xlApp.ActiveSheet.Name 
    def SetSheetName(self, name):
        self.xlApp.ActiveSheet.Name = name 
        
    def AddCharts(self, datalst, charttype = 4):
        #4 xlLine
        import random
        self.newSheet("addCharstmp"+str(random.randint(1000,10000) ))
        self.xlApp.Visible = 1
        row = 0
        for lst in datalst:
            row += 1
            col = 0
            for v in lst:
                col += 1
                self.setCell(row, col, v)
        #self.save()
        #self.xlBook.Charts.ChartType(charttype)
        self.xlBook.Charts.Add()
        self.xlBook.ActiveChart.ChartType = charttype
        
    def do_copyhere(self):
        try:
            self.xlApp.Selection.Copy()
        except Exception, e:
            print "copy chart err: "
            self.printerr(e)
        
        
    def do_copychart(self, sheetname, chartidx):
        try:
            self.activateSheet(sheetname)
            self.xlApp.ActiveSheet.ChartObjects(chartidx).Select()
            self.xlApp.Selection.Copy()
            return True
        except Exception, e:
            print "copy chart err: "
            self.printerr(e)
            return False
            
    def cal_Macro(self, sheetname,marconame, *paras):       
        try:
            self.activateSheet(sheetname) 
            self.xlApp.Run(marconame, *paras)  

        except Exception, e:
            print "calMacro err: "
            self.printerr(e)
            
            
        
        
        
                
        


#-------------------------------I'm split line---------------------         
if __name__ == "__main__":  
    excels = ExcelHelper(r"G:\down\ChrDw\双色球、3D开奖号图表(含奇偶大小).xls")
    sheetnum = excels.getSheetCount()
    for i in range(sheetnum):
        excels.getSheet(i+1)
        sheetname = excels.GetSheetName()
        print sheetname
    
    print(str(excels.getSheetCount()))
    print(str(excels.GetUsedColsNum()))
    print(str(excels.GetUsedRowsNum()))
    print(excels.GetSheetName())
    #excels.SetSheetName("123")
    print(excels.GetSheetName())
    #excels.save("d:\\11.xls")
    excels.close()
    print("It's ok")
    