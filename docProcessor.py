#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jan 22 16:52:27 2019

@author: zhenhao
"""

from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import pandas as pd

filePath='./data/nems_111010.docx'
filePath='./data/sample.docx'
#filePath='./data/017921w14 Pattern Manager 6.0 - Operating User Guide.docx'
#filePath='./data/Unstructured_Assurance of DOT and TBiz Voice on nbn FTTP.docx'    

class docs:
    def __init__(self,text, rank, bold):
        self.text=text
        self.rank=rank
        if bold=='True':
            self.bold=1
        else:
            self.bold=0
        self.length=len(self.text)
        self.child=[]
        self.parent=None
        
    def assignChild(self,ele):
        self.child.append(ele)
    
    def assignParent(self,ele):
        self.parent=ele
        
class docStruct:
    def __init__(self):
        self.parent=docs('parent',0, 1)
        self.store=[self.parent]
        self.prev=self.parent
        self.cur=self.parent
        self.dictStore=None
        
    def findParents(self):
        if self.prev.rank<self.cur.rank or self.prev.bold>self.cur.bold:
            self.cur.assignParent(self.prev)
            self.prev.assignChild(self.cur)
            self.prev=self.cur
        elif self.prev.rank==self.cur.rank:
            self.cur.assignParent(self.prev.parent)
            self.prev.parent.assignChild(self.cur)
            self.prev=self.cur
            
        if self.prev.rank>self.cur.rank:
            self.prev=self.prev.parent
            self.findParents() 
    
    def addNewNode(self, text, rank, bold):
        self.cur=docs(text, rank, bold)
        self.store.append(self.cur)
        self.findParents()
        
    def compileLoop(self,head):
        if len(head.child)==0:
            return head.text
        store=[]
        for i in head.child:
            store.append(self.compileLoop(i))
        
        return {head.text:store}
        
    def compileDict(self):
        store=self.compileLoop(self.parent)
        return store
    
    def processDF(self,df):
        for i in range(len(df)):
            self.addNewNode(df.iloc[i,0], df.iloc[i,7], df.iloc[i,1])
        
        self.dictStore=self.compileDict()

def fontExtractor(paragraph):
    ans=[]
    ans2=[]
    try:
        runs=paragraph.runs[0].font
        ans.append(str(paragraph.text))
        ans.append(str(runs.bold))
        ans.append(str(runs.italic))
        ans.append(str(runs.underline))
        ans.append(str(runs.size))
        ans.append(str(runs.name))
        ans.append(str(paragraph.alignment))
    except:
        t=1
    
    style=paragraph.style.font
    ans2.append(str(paragraph.text))
    ans2.append(str(style.bold))
    ans2.append(str(style.italic))
    ans2.append(str(style.underline))
    ans2.append(str(style.size))
    ans2.append(str(style.name))
    ans2.append(str(paragraph.alignment))
    
    finalAns=[]
    
    for count, rec in enumerate(ans):
        if rec!='None':
            finalAns.append(rec)
        else:
            finalAns.append(ans2[count])
    
    return finalAns

def docProfileCreator(doc):
    df=pd.DataFrame(columns=['text', 'bold', 'italic', 'underline', 'size', 'name', 'alignment'])
    for count, i in enumerate(list(doc.paragraphs)):
        if i.text!='':
            df.loc[count]=fontExtractor(i)
            
    sizes=[int(i) for i in set(df['size'])]
    sizes.sort(reverse=True)
    sizeVal={val:count+1 for count, val in enumerate(sizes)}
    
    df['rank']=[sizeVal[int(val)] for val in df['size']]
        
    return df

def findParents(ele, prev):
    if prev.rank>ele.rank:
        ele.assignParent(prev)
        prev.assignChild(ele)
        return prev
    if prev.rank==ele.rank:
        ele.assignParent(prev.parent)
        prev.parent.assignChild(ele)
        return ele
    if prev.rank<ele.rank:
        findParents(ele, prev.parent)

doc=Document(filePath)
df=docProfileCreator(doc)

docStore=docStruct()
docStore.processDF(df)
docDict=docStore.dictStore

tables=[]
for table in doc.tables:
    rows=[]
    for row in table.rows:
        cells=[]
        try:
            for cell in row.cells:
                cells.append(cell.text)
            rows.append(cells)
        except:
            t=1
    
    tables.append(rows)
    