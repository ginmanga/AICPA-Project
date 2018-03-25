""" Script to read AICPA Word Files"""
# First gather identifying data and place it into a spreadsheet
import os
import docx
import glob

def getText(filename, a):
    """Function gathers text fom docx file"""
    #Will add a variable that takes the list of paragraph numbers within the file
    #And looks for the needed text to gather data
    doc = docx.Document(filename)
    fullText = []
    b = a[2][0]
    print("GETTEXT")
    print(filename)
    print(doc.paragraphs[b].text)
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

def fsttotal(file_path, file_name):
    """Function to find start and total documents"""
    los = ['of', 'DOCUMENTS']
    a = [file_name]
    print("FSTTTTTTTOOOTTAkllll")
    print("File name %s" % (a))
    print("File path  %s" %(file_path))
    file_doc = docx.Document(file_path)
    #print("MAAAADDDEE")
    paras = file_doc.paragraphs
    a.extend(fnd(paras, los))
    return a

def file_loop(path):

def fipath(gvkey, path):
    """Function delivers path to files to open"""
    path = os.path.abspath(path)
    print(os.path.isdir(path))
    if os.path.abspath(path) == False:
        file_name = os.path.splitext(os.path.basename(path))[0]
        # get file name without path or extension
        a = fsttotal(path, file_name)
        return a

    #try:
        #print("HERE?")
        #file_name = os.path.splitext(os.path.basename(path))[0]
        # get file name without path or extension
        #a = fsttotal(path, file_name)
        #return a
    #except:
        #None
    for file in os.listdir(path):
        #Loops through files and folders in path
        #calls fsttotal function
        file_path_a = os.path.join(path, file)
        if os.path.isdir(file_path_a) == True:
            for i in os.listdir(file_path_a):
                file_path_open = os.path.join(file_path_a, i)
                print("FIPATH path %s" % (file_path_open))
                a = fsttotal(file_path_open, os.path.splitext(i)[0])
                getText(file_path_open, a)
                print(a)
        else:
            a = fsttotal(file_path_a, os.path.splitext(file)[0])
            print(a)
    return a