# Introduction
This repository contains code base for Document Intelligence for pdf documents

    1) Converting PDF to docx
    2) Extracting header / text pairs in a csv format
    3) Getting a tree representation of the docx (json format)
    
# General steps
    1) Clone Repo
    2) install requirements.txt

# (1) Converting PDF to docx

     1) run "python fileConverter.py -f <input file path> -o <output file path> -t <filetype. docx by default>
     
     the pdf file at "input file path" will be converted to docx and stored at "output file path"

# (2) Extracting header / text pairs in a csv format

     1) run "python docProcessor.py -f <input file path> -o <output file path>
     
     the docx file at "input file path" will be processed and the output csv with header and text pairs will be saved at "output file path"

# (3) Getting a tree representation of the docx (json format)
    1) in your own script:
        a) import docProcessor
        b) initialise class and call generateHeaderTextCSV(<filePath>)






