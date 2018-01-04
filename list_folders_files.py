
import os
import docx2txt
path = "C:\\Users\\aelsalla\\Documents\\Valeo Documents\\Official & Mgmt\\Screening CVS\\DAS\\Screening"

'''
try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile


"""
Module that extract text from MS XML Word document (.docx).
(Inspired by python-docx <https://github.com/mikemaccana/python-docx>)
"""




def read_doc(path):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)

    paragraphs = []
    for paragraph in tree.getiterator(PARA):
        texts = [node.text
                 for node in paragraph.getiterator(TEXT)
                 if node.text]
        if texts:
            paragraphs.append(''.join(texts))

    return '\n\n'.join(paragraphs)
'''
import sys
import os
import comtypes.client

def convert_doc2docx(full_file_name):
    in_file = full_file_name
    out_file = full_file_name.replace(".doc", ".docx")  # name of output file added to the current working directory
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)  # name of input file
    doc.SaveAs(out_file, FileFormat=16)  # output file format to Office word Xml default (code=16)
    doc.Close()
    word.Quit()
    return out_file

import docx2txt
def read_docx(full_file_name):

    fullText = docx2txt.process(full_file_name)
    text = docx2txt.process(full_file_name)

    return fullText

import PyPDF2
def read_pdf(full_file_name):

    pdfFileObj = open(full_file_name, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    fullText = ""
    for page in range(pdfReader.numPages):
        fullText = fullText + pdfReader.getPage(page).extractText()
    return fullText

for root, dirs, files in os.walk(path):
    if(dirs == []):
        #print(root)
        label = root.split("\\")[-1]
        print(label)
        for file in files:

           #f = open(root + "\\" + file, 'rb')
            print(file)
            full_file_name = root + "\\" + file

            if(os.path.splitext(file)[1] == ".docx"):
                text = read_docx(full_file_name)
            elif(os.path.splitext(file)[1] == ".doc"):
                full_file_name_doc = convert_doc2docx(full_file_name)
                text = read_docx(full_file_name_doc)

            elif(os.path.splitext(file)[1]==".pdf"):
                text = read_pdf(full_file_name)
            else:
                print("Unsupported file", file)






