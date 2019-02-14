#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jan 22 13:55:44 2019

@author: zhenhao
"""

import subprocess
from docx import Document

lowriter='/usr/bin/soffice'
lowriter= 'soffice'

def extractName(docPath, etx, storePath, orgEtx='.pdf'):
    tem=docPath.split('/')
    tem=tem[len(tem)-1]
    tem=storePath+tem[:-len(orgEtx)]+'.'+etx
    
    return tem

def convertPDF(docPath, storePath, etx='doc'):    
    record='{0} --infilter="writer_pdf_import" --convert-to {1} --outdir "{2}" "{3}"'.format(lowriter, etx, storePath, docPath)
    print(record)
    newDocPath=extractName(docPath, etx, storePath)
    subprocess.call(record, shell=True)
    
    return newDocPath
    
#docdir='documents/test.pdf'
#outdir='store/'
#
#newPath=convertPDF(docdir, outdir, 'odt')


filePath='./data/nems_111010.docx'
doc=Document(filePath)