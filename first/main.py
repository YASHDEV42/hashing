import pandas as pd
import camelot
import PyPDF2
from docx2pdf import convert

import json

convert("test.docx", "test.pdf")



paragraphs = camelot.read_pdf('test.pdf', flavor='stream')
paragraphs[0].to_json('simple.json')



with open('test.pdf', 'rb') as pdf_file:
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    text_content = ""
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        text_content += page.extract_text()

lines = text_content.splitlines()

if lines:
    df = pd.DataFrame(lines, columns=['Text'])


myCode = ""
if "حفظه للاه" in df.iloc[0]['Text']:

    
    myCode = "1"

    if "مدير إدارة المخاطر" in df.iloc[0]['Text']:
        myCode += "-PR"

    if "مشرف الوحدة الخامسة" in df.iloc[df.shape[0] - 1]['Text']:
        myCode += "-VIC1"


contentCounter = 0

tables = camelot.read_pdf('test.pdf', pages='1-end')
if tables:
    contentCounter += 1
    myCode += "-SUB" + str(contentCounter)
    myCode += "-T"
    tables[0].to_json('foo.json')

    data = json.load(open('foo.json'))
    myStr = ""
    for item in data:
        for key, value in item.items():
            if key == "2":
                myStr += value
            else:
                myStr += value + "."
        myStr += "-"    
    
    myCode += str(len(data[0])) + "." + str(len(data))
    
myCode += "-"  + myStr + "SUB" + str(contentCounter+1)
print(myCode)
filename = "output.txt"
with open(filename, "w") as f:
    f.write(myCode)