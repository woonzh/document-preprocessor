#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Jan 22 13:55:44 2019

@author: zhenhao
"""

import subprocess

lowriter='/usr/bin/soffice'
lowriter= 'soffice'

docdir='documents/test.pdf'
outdir='store/'

record='{0} --infilter="writer_pdf_import" --convert-to doc "{1}"'.format(lowriter, docdir)

subprocess.call(record, shell=True)