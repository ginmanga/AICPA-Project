""" Script to read AICPA Word Files"""
# First gather identifying data and place it into a spreadsheet
import os
import docx
#from docx import Document
file_experiment = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\Annual_Reports_-_Corporate_(AICPA)__1972-1982011-05-07_23-05.docx'
file_experiment = os.path.abspath(file_experiment)
file_test = docx.Document(file_experiment)
sections = file_test.sections
paras = file_test.paragraphs

count = 0


def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)
#file_text = getText(file_experiment)
#print(file_text[0:500])
count = 0
for i in file_test.paragraphs:
    if i.text != '' and i.text.isspace() == False:
        print(i.text)
        print(count)
    if count>200:
        break
    count += 1
#for i in file_test.sections:
    #if i.text != '' and i.text.isspace() == False:
        #print(i.text)
        #print(count)
        #break
    #count += 1