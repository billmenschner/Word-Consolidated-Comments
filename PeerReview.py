#! Python3

# PeerReview.py -- Pulls MS Word comments from docx files and copies them into
# consolidated file

import os, shutil, tkinter as tk, glob, re, zipfile, docx, pandas as pd
from tkinter import filedialog
from pathlib import Path
from lxml import etree
from roman import toRoman

#Start by having the user select the directory where the files are located
root = tk.Tk()
root.withdraw()

file_path = filedialog.askdirectory()

#Check the folder for docx files
data_folder = Path(file_path).glob('*.docx')
data_folder_count = len(list(data_folder))

if data_folder_count == 0:
    print("This directory doesn't have any docx files. Select a different directory.")
 
    
#Check for the folder that will contain the zip files created from docx files. 
#If it doesn't exist, create folder.
copied_folder = f'{file_path}/CopiedZip'
if data_folder_count > 0:
    if not os.path.exists(copied_folder):
        os.makedirs(copied_folder)   
        
#Copy the docx files to the CopiedZip folder and save them as a zip file 
#in that folder
name_pattern = re.compile(r"""^(.*?) #All the text before the extension
    (.docx)    #The extension
    """, re.VERBOSE)

for current_name in os.listdir(file_path):
    working_name = name_pattern.search(current_name)
    if working_name == None:
        continue
    doc_name = working_name.group(1)
    new_file_name = f'{doc_name}.zip'
    shutil.copy(f'{file_path}/{current_name}', f'{copied_folder}/{new_file_name}')
    
#Create folder where files will be extracted.
    
os.makedirs(f'{copied_folder}/Extracted')
extracted_folder = f'{copied_folder}/Extracted'

#Create folders based on zip file names and extract the documents and comments
#files from the zip files into the new folders.
pulled_text = {}
pulled_comments = {}
namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
comment_number = 0


for file in os.listdir(copied_folder):
    if '.zip' in file:
        os.makedirs(f'{extracted_folder}/{os.path.splitext(file)[0]}')
        new_folder = f'{extracted_folder}/{os.path.splitext(file)[0]}'
        extracted_file = zipfile.ZipFile(f'{copied_folder}/{file}')
        document = extracted_file.extract('word/document.xml', new_folder)
        comments = extracted_file.extract('word/comments.xml', new_folder)

        #Parse both the documents and comments documents and create the dictionary
        #that will hold all of the information taken from the documents.
        
        document_root = etree.parse(document)
        comments_root = etree.parse(comments)
        
        #Find page offset so that page number can be approximated.

        introduction = 0
        current_page = 0
        page_offset = 0
        
        for element in document_root.iter():
            if f'{namespace}lastRenderedPageBreak' == element.tag:
                current_page += 1
            elif f'{namespace}t' == element.tag:
                if element.text == 'Introduction':
                    introduction += 1
                    if introduction == 2:
                        page_offset = current_page
                        break

        #Extract text from document based on Comment ID
        
        text_commentId = ''
        document_text = ''
        page_number = 0
        
        for element in document_root.iter():
            if f'{namespace}t' == element.tag:
                document_text += element.text
            elif f'{namespace}commentRangeStart' == element.tag:
                text_commentId = element.get(f'{namespace}id')
                document_text = ''
            elif f'{namespace}lastRenderedPageBreak' == element.tag:
                page_number += 1
            elif f'{namespace}commentRangeEnd' == element.tag:
                if page_number < page_offset:
                    pulled_text[text_commentId] = (document_text, 
                               toRoman(page_number), page_number)
                else:
                    pulled_text[text_commentId] = (document_text, 
                               page_number - page_offset + 1, page_number)

        
        #Extract Comment ID, comment, and author from comments.xml. Start 
        #running count that gets increased with each comment tag.
        
        commentId = ''
        comment_text = ''
        author = ''
        
        for element in comments_root.iter():
            if f'{namespace}comment' == element.tag:
                if comment_number == 0:
                    commentId = element.get(f'{namespace}id')
                    author = element.get(f'{namespace}author')
                    comment_number += 1
                else:
                    try:
                        pulled_comments[comment_number] = (commentId, author, 
                                       comment_text, pulled_text[commentId][0], 
                                       pulled_text[commentId][1], pulled_text[commentId][2])
                    except KeyError:
                        pulled_comments[comment_number] = (commentId, author, 
                                       comment_text, 
                                       "NOTE: THIS COMMENT WAS NOT ATTACHED TO ANY TEXT",
                                       '', '')
                    comment_number +=1
                    commentId = element.get(f'{namespace}id')
                    author = element.get(f'{namespace}author')
                    comment_text = ''
            elif f'{namespace}t' == element.tag:
                comment_text += element.text      
                
        #Create the Word document
        
        df = pd.DataFrame.from_dict(pulled_comments, orient = 'index', columns = ['comment_ID', 'Author', 'Comment', 'Document Text', 'Document Page/Actual Page', 'Adjudication']).drop('comment_ID', axis=1)
        doc = docx.Document()
        
        #Create a Table with an extra row for headings and an extra column for
        #adjudication.
        
        comment_table = doc.add_table(df.shape[0]+1, df.shape[1]+1, doc.styles['Table Grid'])
        
        #Add headers
        
        for j in range (df.shape[-1]):
            comment_table.cell(0,j).text = df.columns[j]

        #Add comments to the table
        
        for i in range(df.shape[0]):
            for j in range(df.shape[-1]):
                comment_table.cell(i+1,j).text = str(df.values[i,j])
                
        doc.save('test.docx')


