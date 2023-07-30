from docx.document import Document
from docx import Document
from author import Authors, CorrespondingAuthor
from manuscript import ManuscriptInfo
from typing import List

#specific to extracting information from word documents
import os
import zipfile
#other tools useful in extracting the information from our document
import re
#to pretty print our xml:
import xml.dom.minidom
import pandas as pd
import openpyxl

from tkinter import Tk
from tkinter.filedialog import askdirectory
from flask import Flask, jsonify, request

# document = Document("files\SCE3-2023-079_Corrected.docx")
# files\SCE3-2023-094_Corrected.docx
# files\SCE3-2023-079_Corrected.docx
if os.environ.get('DISPLAY','') == '':
    os.environ.__setitem__('DISPLAY', ':0.0')

class CETExtraction():
    authors: Authors
    manuscript_info: ManuscriptInfo
    paper_id : str

    def __init__(self, filename: str):
        document = Document(filename)
        self.paper_id = self._get_paper_id(filename = filename)
        print(self.paper_id)
        for p, paragraph in enumerate(document.paragraphs):
            if paragraph.style.name == 'CET Authors' or p == 1:
                # self.authors = self._get_authors_names(paragraph)
                self.authors = self._get_authors_names_2(paragraph)
                break

        for p, paragraph in enumerate(document.paragraphs):
            if p == 0:
                self.manuscript_info = self._get_manuscript_info(filename, paragraph)
                break

    
    def _get_paper_id(self, filename: str):
        # return filename.split('-')[-1].split('_')[0] + '.pdf' # SCE
        return filename.split('/')[-1].split('_')[1].split('_')[0] + '.pdf'

    def _get_manuscript_info(self, filename, paragraph):
        page_number = int(self._get_page_number(filename = filename))
        paper_title = paragraph.text
        return ManuscriptInfo(page_no = page_number, paper_title = paper_title)

    def _get_page_number(self, filename):
        try:
            document = zipfile.ZipFile(filename)
            dxml = document.read('docProps/app.xml')
            uglyXml = xml.dom.minidom.parseString(dxml)
            page = uglyXml.getElementsByTagName('Pages')[0].childNodes[0].nodeValue
        except Exception as e:
            # raise
            page = 0
            pass
        # print("Word Page count: " + page)
        return page

    def _get_authors_names_2(self, paragraph):
        authors = []
        corresponding_author = [paragraph.text.strip().split('*')[0]]
        corresponding_author = [author.split(',')[0] for author in corresponding_author]
        authors = [word.strip().split('*')[0] for word in paragraph.text.split(',') if len(word.strip()) > 1] # Split the text, and avoid getting the superscripts, and split from '*'

        is_contain_superscripts = False
        for r, run in enumerate(paragraph.runs):
            if run.font.superscript is True:
                is_contain_superscripts = True
                break
        
        if is_contain_superscripts:
            authors = [author[:-1] for author in authors]
            corresponding_author = [author[:-1] for author in corresponding_author]
            pass
        
        print(authors)
        return Authors(author_list = authors, corresponding_author = corresponding_author)
    
    def _get_authors_names(self, paragraph):
    # ---------------------------------------
    # Author names
    #----------------------------------------
        authors = []
        new_author = True
        
        for r, run in enumerate(paragraph.runs):
            # authors = [word.strip() for word in paragraph.text.split(',')]
            if (
                (run.font.superscript is None and '*' != run.text.strip()) and # Avoid superscript and only '*'
                '' != run.text.strip() or # If is not empty texts
                ('’' == run.text.strip() or '.' == run.text.strip() or '-' == run.text.strip()) # If only the special characters in authors name, can be included
            ): 
                if ("," in run.text):
                    check = run.text.split(',')
                    for text in run.text.split(','):
                        if text and len(text) > 1:
                            if ( 
                                not new_author
                            ):
                                authors[-1] += ' ' + text.strip().split('*')[0] if ('’' != run.text.strip() and '.' != run.text.strip() and '-' != run.text.strip()) else text.strip().split('*')[0]
                            else:
                                authors += [text.strip().split('*')[0].strip()]
                                new_author = False
                        
                        # Condition, if run.text is [', xxx' or 'yyy , xxx'] or just ','
                        new_author = True if (text != run.text.split(',')[-1] or '' == run.text.split(',')[-1]) else new_author

                else:
                    if ( 
                        not new_author
                    ):
                        authors[-1] += ' ' + run.text.strip().split('*')[0] if ('’' != run.text.strip() and '.' != run.text.strip() and '-' != run.text.strip()) else run.text.strip().split('*')[0]
                    else:
                        authors += [run.text.strip().split('*')[0].strip()]
                        new_author = False
            
            elif '' == run.text.strip():  # Ignore empty texts
                continue
            else:
                new_author = True

        print(authors)       
        return Authors(author_list = authors)
    
class CETManuscripts():
    all_info: List[CETExtraction]

    def __init__(self, file_list: List[str], file_path: str):
        self.all_info = []
        for file_name in file_list:
            filepath = f"{file_path}/{file_name}"
            self.all_info.append(CETExtraction(filename = filepath)) 
    
    def write_to_excel(self, file_path: str):
        rows_of_data_in_excel = []
        for manuscript_info in self.all_info:
            # (1): paper title
            # (2): page count
            # (3): paper id.pdf
            # (4): number of authors
            # (5): first name of first author
            # (6): last name of first author

            row = [
                manuscript_info.manuscript_info.paper_title,
                manuscript_info.manuscript_info.page_no,
                manuscript_info.paper_id,
                manuscript_info.authors.no_of_authors
            ]
            for author_ind in range(0, manuscript_info.authors.no_of_authors):
                row.append(manuscript_info.authors.first_name[author_ind])
                row.append(manuscript_info.authors.last_name[author_ind])

            rows_of_data_in_excel.append(row)
        
        df = pd.DataFrame(rows_of_data_in_excel)
        # df.to_excel(f'{file_path}//PRES23_CET_Info.xlsx', sheet_name = 'LAVOLI')
        df.to_excel(f'{file_path}//PRES23_CET_Info.xlsx', sheet_name = 'LAVORI')

# Creating a Web App
app = Flask(__name__)

@app.route('/', methods = ['GET'])
def get_CET_info():
    path = askdirectory(title='Select Folder') # shows dialog box and return the path
    # file_list = [file for file in os.listdir(path) if ('docx' in file and '$' not in file)]
    file_list = [file for file in os.listdir(path) if ('docx' in file and '$' not in file and 'PRES23' in file)]
    # file_list = ["+PRES23_0184_M_v_03_Author.docx"]
    # file_list = ["+PRES23_0340_rev_M_v2_Review.docx"]
    # file_list = ["+PRES23_0002_rev_M_v2_Review.docx"]
    CET_manuscripts = CETManuscripts(file_list = file_list, file_path = path)
    CET_manuscripts.write_to_excel(path)

    response = {
        'message': f"Success! All CET info in the folder is extracted and saved to {path}/PRES23_CET_Info.xlsx"
    }

    return jsonify(response), 200

# Running the app
# app.run(host = '0.0.0.0', port = 5000)


if __name__ == "__main__":
    app.run(host = '0.0.0.0', debug = False)
    # path = askdirectory(title='Select Folder') # shows dialog box and return the path
    # # file_list = [file for file in os.listdir(path) if ('docx' in file and '$' not in file)]
    # file_list = [file for file in os.listdir(path) if ('docx' in file and '$' not in file and 'PRES23' in file)]
    # # file_list = ["+PRES23_0184_M_v_03_Author.docx"]
    # # file_list = ["+PRES23_0340_rev_M_v2_Review.docx"]
    # # file_list = ["+PRES23_0002_rev_M_v2_Review.docx"]
    # CET_manuscripts = CETManuscripts(file_list = file_list, file_path = path)
    # CET_manuscripts.write_to_excel(path)

    pass
