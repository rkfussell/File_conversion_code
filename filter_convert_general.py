# -*- coding: utf-8 -*-
"""
Created on Wed Aug  2 14:04:37 2023

@author: rkf33
"""
import os
import pandas as pd
import secrets
import shutil
import win32com.client
import docx
import nltk
import fitz


def process_files(dirname):
    #save list of filenames
    folders = []
    filenames = []
    for path, subdirs, files in os.walk(dirname):
        for subdir in subdirs:
            folders.append(os.path.join(dirname, subdir))
    for folder in folders:
        for path, subdirs, files in os.walk(folder):
            for name in files:
                filenames.append(os.path.join(folder, name))  
    #filter out non pdf and non word documents
    extensions = [".pdf", ".docx",".doc"]
    filenames = [f for f in filenames if os.path.splitext(f)[1] in extensions]
    new_filenames = [os.path.splitext(f)[0] + "_2" + os.path.splitext(f)[1] for f in filenames ]
    #print(filenames)
    
    #n = len(filenames)

    

def pdf_to_docx(dirname):
    
    word = win32com.client.Dispatch("Word.application")
    
    files = []
    for file in os.listdir(dirname):
        if os.path.splitext(file)[1] == ".pdf":
            files.append(os.path.join(dirname, file))
    for file in files:
        wordDoc = word.Documents.Open(file, False, False, False)
        wordDoc.SaveAs2(os.path.splitext(file)[0])
        wordDoc.Close()
    print(files)

def docx_to_xlsx(dirname):
    for fileloc in os.listdir(dirname):
        if os.path.splitext(fileloc)[1] == ".docx":
            print(fileloc)
            with open(os.path.join(dirname,fileloc),"rb") as file:
                document = docx.Document(file)
                
            raw_text = []
            raw_text_context = []
            for paragraph in document.paragraphs:
                if paragraph.text != "" and paragraph.text != " ":
                    for sentence in nltk.sent_tokenize(paragraph.text):
                        raw_text.append(sentence)
                        raw_text_context.append(paragraph.text)
            file.close()
            with pd.ExcelWriter(os.path.splitext(fileloc)[0] + '.xlsx') as writer:
                pd.DataFrame({'Sentences': raw_text}).to_excel(
                        writer,
                        sheet_name='Sentence Data')
        


main_filepath = r"C:\Users\rkf33\Documents\File_conversion_code\cleaned"

process_files(main_filepath)
pdf_to_docx(main_filepath)
docx_to_xlsx(main_filepath)



