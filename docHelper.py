# !/usr/bin/
# coding=gbk
'''
Created on 2013年10月10日

@summary: 

@author: huxiufeng
'''
from win32com.client import Dispatch
import os

class docHelper:
    '''
    Some convenience methods for Excel documents accessed
    through COM.
    '''
    def __init__(self, visible=False ):
        self.wdApp = Dispatch('Word.Application')
        self.wdApp.Visible = visible
    
    def open(self, filename=None):
        if filename:
            self.filename=filename
            if os.path.exists(self.filename):
                try:
                    self.wdApp.Documents.Open(self.filename)
                except Exception , e:
                    print "open err： "
                    self.printerr(e)
                    return False
            else:
                self.wdApp.Documents.Add()
                self.save(self.filename,True)
        else:
            self.wdApp.Documents.Add()
            self.filename='Untitle'
        return True
          
    def visible(self, visible=True):
        self.wdApp.Visible = visible
    
    def find(self, text, MatchWildcards=False):
        '''
        Find the string
        '''
        find = self.wdApp.Selection.Find
        find.ClearFormatting()
        find.Execute(text, False, False, MatchWildcards, False, False, True, 0)
        return self.wdApp.Selection.Text
          
    def replaceAll(self, oldStr, newStr):
        '''
        Find the oldStr and replace with the newStr.
        '''
        find = self.wdApp.Selection.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        find.Execute(oldStr, False, False, False, False, False, True, 1, True, newStr, 2)   
    
    def updateToc(self):
        for tocitem in self.wdApp.ActiveDocument.TablesOfContents:
            tocitem.Update()
    
    def save(self, filename = None, delete_existing=True):
        '''
        Save the active document
        '''
        if filename is None:
            self.wdApp.ActiveDocument.Save()
        else:
            if delete_existing and os.path.exists(filename):
                os.remove(filename)
            self.wdApp.ActiveDocument.SaveAs(FileName=filename)
            
    def printerr(self, e):
        for v in e:
            #print type (v)
            if type(v) == type([]) or type(v) == type((1,2)):
                for vv in v:
                    print "\t\t", vv
            else:
                print "\t", v 
        
        
    def close(self):
        '''
        Close the active workbook.
        '''    
        try:
            #self.wdApp.ActiveDocument.Close()
            self.wdApp.Documents.Close()
        except Exception, e:
            print "close err： "
            self.printerr(e)
        del self.wdApp
          
    def quit(self):
        '''
        Quit Word
        '''
        return self.wdApp.Quit()
    
    def AddTable(self, lsttable):
        row = len(lsttable)
        col = 0
        for lst in lsttable:
            if len(lst) > col:
                col = len(lst)
        #print row, col
        
        tb = self.wdApp.ActiveDocument.Tables.Add(self.wdApp.Selection.Range, row, col)
        
        tb.Rows[0].Cells[0].Range.Text = "Table0"
        
        '''
        for l in range(row-2):
            self.wdApp.ActiveDocument.Tables[tnum].Rows.Add()
        
        for l in range(col-2):
            self.wdApp.ActiveDocument.Tables[tnum].Columns.Add()
        '''
        #self.wdApp.ActiveDocument.Tables.Add(self.wdApp.Selection.Range, row+1, col+1)
        
        i = 0
        for lst in lsttable:
            j = 0
            for l in lst :  
                tb.Rows[i].Cells[j].Range.Text = str(l)
                j = j+1
            i = i+1
        
        tb.AutoFormat(ApplyBorders=True)
        
    def goTop(self):
        '''return to the head of file'''
        #Selection.HomeKey unit:=wdStory
        #self.wdApp.Selection.HomeKey(wdStory)
        self.wdApp.ActiveDocument.GoTo(What=11)
        self.wdApp.ActiveDocument.Select()
#         self.wdApp.ActiveDocument.Range(
#                     Start = 0, 
#                     End = 0 
#                     ).Select()
        
    def goEnd(self):
        self.wdApp.ActiveDocument.Range(
                    Start = self.wdApp.ActiveDocument.Content.End-1, 
                    End = self.wdApp.ActiveDocument.Content.End-1 
                    ).Select()
        
        
    def AddPic(self,oldstr,filename, height = 150, width = 250):
        self.goTop()
        strs = self.find(oldstr)
        if strs == oldstr and os.path.isfile(filename):
            shp = self.wdApp.Selection.InlineShapes.AddPicture(filename)
            shp.Height = height
            shp.Width =width
    
    def do_paste(self,tag = None):
        try:
            if tag is None :
                self.wdApp.Selection.Paste()
            else:
                self.goTop()
                if self.find(tag) == tag:
                    self.wdApp.Selection.Paste()
                else:
                    print "do not find tag: ", tag, " in the doc"
                return True     
        except Exception, e:
            print "paste err: "
            self.printerr(e)
            return False
            
    

#----------------------It is a split line--------------------------------------

def main():
    docpath = r'res/doc_demo.doc'
    abspath = os.path.abspath(docpath)
    doc = docHelper(True)
    if not doc.open(abspath):
        print "open ", abspath, " error"
        return
    datalst = [["col1", "col2", "col3"], [1,2,3], [3,4,5],[5,6,7],[9,8,7],[6,5,4]]
    
#     doc.do_paste()
#     doc.goEnd()
#     doc.AddTable(datalst)
#     doc.goTop()
#     doc.do_paste()
#     doc.goEnd()
    doc.do_paste()
    doc.save()
    #doc.close()
    
#----------------------It is a split line--------------------------------------

if __name__ == "__main__":
    main()
    print "It's ok"
