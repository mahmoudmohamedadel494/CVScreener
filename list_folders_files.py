
import docx2txt
import os
import comtypes.client
import PyPDF2
import pandas as pd




def convert_doc2docx(full_file_name):
    in_file = full_file_name
    out_file = full_file_name.replace(".doc", ".docx")  # name of output file added to the current working directory
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)  # name of input file
    doc.SaveAs(out_file, FileFormat=16)  # output file format to Office word Xml default (code=16)
    doc.Close()
    word.Quit()
    return out_file

def read_docx(full_file_name):

    fullText = docx2txt.process(full_file_name)
    text = docx2txt.process(full_file_name)

    return fullText


def read_pdf(full_file_name):

    pdfFileObj = open(full_file_name, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    fullText = ""
    for page in range(pdfReader.numPages):
        fullText = fullText + pdfReader.getPage(page).extractText()
    return fullText

def load_data(path):
    d = []
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
                    os.remove(full_file_name_doc)

                elif(os.path.splitext(file)[1]==".pdf"):
                    text = read_pdf(full_file_name)
                else:
                    print("Unsupported file", file)

                d.append({'text': text, 'label': label , 'file': file})


    df = pd.DataFrame().from_dict(d)
    return df

path = "C:\\Users\\aelsalla\\Documents\\Valeo Documents\\Official & Mgmt\\Screening CVS\\DAS\\Screening"
df = load_data(path=path)
print(df.head(3))

