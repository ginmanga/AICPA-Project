""" Script to read AICPA Word Files"""
# First gather identifying data and place it into a spreadsheet
import os
import docx
#from docx import Document
file_experiment = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\Annual_Reports_-_Corporate_(AICPA)__1972-1982011-05-07_23-05.docx'
file_experiment = os.path.abspath(file_experiment)
file_test = docx.Document(file_experiment)
#sections = file_test.sections
paras = file_test.paragraphs

directory_a = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\NO GVKEY'
#print(os.listdir(directory_a))

print(paras[9].text)
print(any(char.isdigit() for char in paras[9].text))

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

def fnd(paragraphs, terms):
    """Given a string of characters find paragraph numbers of each case"""
    #For AICPA files, look for number of number DOCUMENT
    count_par = 0
    count_doc = 0
    list_paras = []
    for i in paragraphs:
        fc = terms[0] in i.text
        sc = terms[1] in i.text
        dc = any(char.isdigit() for char in i.text)
        c_list = [fc, sc, dc]

        #if fc == True and sc == True:
        if all(cond == True for cond in c_list):
            #print(i.text)
            #print(count_par)
            list_paras.append(count_par)
            count_doc += 1
            #print(count_doc)
        count_par += 1
    return count_doc, list_paras

a = fnd(paras, los)

#b = [0 , 0 ,0]
#bb = [0 ,1, 0]

#print(all(i == 0 for i in bb))
#print(a[0])
#print(a[1])

directory = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\NO GVKEY'
#print(os.listdir(directory_a))
los = ['of', 'DOCUMENTS']



def parse_AICPA(gvkey, path):
    for file in os.listdir(path):
        file_path = os.path.join(directory_a, file)
        print(file)
        file_doc = docx.Document(file_path)
        paras = file_doc.paragraphs
        a = fnd(paras, los)
        print(a[0])

parse_AICPA(0, directory)

#for i in file_test.sections:
    #if i.text != '' and i.text.isspace() == False:
        #print(i.text)
        #print(count)
        #break
    #count += 1