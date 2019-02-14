# Introduction
This repository contains code base for Document Intelligence for pdf documents

  - Data Preparation
  - Model building and training
  - Final pipeline

# (1) Data preparation

  - Convert pdf to docx format
  - Preprocess the docx file and tag relevant paragraphs
  - Label the paragraghs for building NLP model

Further enhancement would be:
  - Labelling and storing Tables contents, Figures
  - List and Paragraphs 


# (2) Installation Steps

- Step 1 : Run the requirements.txt to install the required libraries
- Step 2 : Run the data-preparation.py with the input docx file to be extracted
- Step 3 : Extracted and tagged contents will be placed in the /out_file (csv file format)

```sh
$ pip install -r requirements.txt
$ python data-preparation.py [path-to-ip-file-name]
$ cd out_file
```







