#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jan 22 13:55:44 2019

@author: zhenhao
"""

import subprocess
import argparse

lowriter='/usr/bin/soffice'
lowriter= 'soffice'

def extractName(docPath, etx, storePath, orgEtx='.pdf'):
    tem=docPath.split('/')
    tem=tem[len(tem)-1]
    tem=storePath+tem[:-len(orgEtx)]+'.'+etx
    
    return tem

def convertPDF(docPath, storePath, etx='docx'):    
    record='{0} --infilter="writer_pdf_import" --convert-to {1} --outdir "{2}" "{3}"'.format(lowriter, etx, storePath, docPath)
    print(record)
    newDocPath=extractName(docPath, etx, storePath)
    subprocess.call(record, shell=True)
    
    return newDocPath
    
#docdir='documents/test.pdf'
#outdir='store/'
#
#newPath=convertPDF(docdir, outdir, 'odt')

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process docx files')
    parser.add_argument("-f", dest='filePath', action="store", help='input file path')
    parser.add_argument("-o", dest='outFilePath', action="store", help='input file path')
    
    parser.add_argument("-t", dest='fileType', default='docx',action="store", help='input file path')
    args = parser.parse_args()
    
    newPath=convertPDF(args.filePath, args.outFilePath, args.fileType)
    print('file converted and saved to %s'%(newPath))