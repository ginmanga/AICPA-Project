""" Script to read AICPA Word Files"""
# First gather identifying data and place it into a spreadsheet
import os
import docx
import glob
def getText(filename):
    """Function gathers text fom docx file"""
    #Will add a variable that takes the list of paragraph numbers within the file
    #And looks for the needed text to gather data
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

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
        if all(cond == True for cond in c_list):
            list_paras.append(count_par)
            count_doc += 1
        count_par += 1
    return count_doc, list_paras

directory = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\Annual_Reports_-_Corporate_(AICPA)__1972-1982011-05-07_23-05.docx'
directory = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\GVKEY'
directory = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\NO GVKEY'
directory = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\Annual_Reports_-_Corporate_(AICPA)__1972-1982011-05-07_23-05.docx'
los = ['of', 'DOCUMENTS']

def fsttotal(file_path, file_name):
    """Function to find start and total documents"""
    a = [file_name]
    print(a)
    file_doc = docx.Document(file_path)
    print("MAAAADDDEE")
    paras = file_doc.paragraphs
    a.extend(fnd(paras, los))
    return a

def fipath(gvkey, path):
    """Function delivers path to files to open"""
    path = os.path.abspath(path)
    try:
        file_name = os.path.splitext(os.path.basename(path))[0]
        # get file name without path or extension
        a = fsttotal(path, file_name)
        return a
    except:
        None
    for file in os.listdir(path):
        #Loops through files and folders in path
        #calls fsttotal function
        file_path_a = os.path.join(path, file)
        if os.path.isdir(file_path_a) == True:
            for i in os.listdir(file_path_a):
                file_path_open = os.path.join(file_path_a, i)
                a = fsttotal(file_path_open, os.path.splitext(i)[0])
        else:
            a = fsttotal(file_path_a, os.path.splitext(file)[0])
    return a

print(fipath(0, directory))