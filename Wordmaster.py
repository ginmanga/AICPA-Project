""" Script to read AICPA Word Files"""
# First gather identifying data and place it into a spreadsheet
import os
import docx
import glob


def getText(filename, file_details):
    """Function gathers text fom docx file"""
    #Will add a variable that takes the list of paragraph numbers within the file
    #And looks for the needed text to gather data
    doc = docx.Document(filename)
    fullText = []
    para = doc.paragraphs
    print("Have entered GETEXT")
    #print(len(para))
    #print(para[9].text)
    for i in file_details[2]:
        print(para[i].text)
        #print(doc.paragraphs.text[i])
    for para in doc.paragraphs:
        #print(para.text)
        fullText.append(para.text)
    return '\n'.join(fullText)


#def par_loop(file_path, file_details):
    #"""Function that loops through paragraph numbers
    #and calls getText for each starting paragraph"""
    #print("Have entered par_loop")
    #for i in file_details[2]:
        #getText(file_path, i,15)

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
    b = []
    a = [file_name]
    file_doc = docx.Document(file_path)
    paras = file_doc.paragraphs
    #print(len(paras))
    a.extend(fnd(paras, los))
    return a


def file_loop(path):
    """This function is called by fipath when the path
    given is a folder and not a file"""
    for file in os.listdir(path):
        #Loops through files and folders in path
        #calls fsttotal function
        file_path_a = os.path.join(path, file)
        if os.path.isdir(file_path_a) == True:
            for i in os.listdir(file_path_a):
                file_path_open = os.path.join(file_path_a, i)
                a = fsttotal(file_path_open, os.path.splitext(i)[0])
                #par_loop(file_path_open,a)
                getText(file_path_open, a)
                #print(a)
        else:
            a = fsttotal(file_path_a, os.path.splitext(file)[0])
            #print(a)
    return a


def fipath(gvkey, path):
    """Function delivers path to files to open"""
    path = os.path.abspath(path)
    #print(os.path.isdir(path))
    if os.path.isdir(path) == False:
        file_name = os.path.splitext(os.path.basename(path))[0]
        # get file name without path or extension
        a = fsttotal(path, file_name)
        return a
    else:
        return file_loop(path)
