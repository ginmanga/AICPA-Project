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

#print(paras[9].text)

#print(paras[9].text[0])

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
#for i in file_test.paragraphs:
for i in paras:
    if i.text != '' and i.text.isspace() == False:
        None
        #print(i.text)
        #print(count)
    if count>200:
        break
    count += 1

los = ['of', 'DOCUMENTS']

#print(terms[0] in paras[9].text)
def fnd(paragraphs, terms):
    """Given a string of characters find paragraph numbers of each case"""
    #For AICPA files, look for number of number DOCUMENT
    for i in paragraphs:
        fc = terms[0] in i.text
        sc = terms[1] in i.text
        if fc == True and sc == True:
            print(i.text)

fnd(paras, los)

#for i in file_test.sections:
    #if i.text != '' and i.text.isspace() == False:
        #print(i.text)
        #print(count)
        #break
    #count += 1