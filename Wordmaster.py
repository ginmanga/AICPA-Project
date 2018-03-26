""" Script to read AICPA Word Files
Not meant to read other types of files
Will have to create other functions for that"""
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
    for i in file_details[3]:
        newText = []
        for j in range(i, i+29):
           (newText.append(para[j].text) if para[j].text != '' else None)
        fullText.append(newText)
    return fullText

def parseText(num_docs, text):
    """Parse the text gotten from geText"""
    id_data = []
    #print(num_docs)
    for i in text:
        print(i[1])
        #print(i[1].find(a))
        #print("NNNNNNN")
        #name = company_name(i[1])
        name = if_find("COMPANY NAME:", i[1])
        #print(i[2:4])
        sic = sic_code(i[2:4])
        #print(sic)


#def company_name(text): LOOKS LIKE DO NOT NEED A FUNCTION FOR THIS
    #"""Receives raw data and returns the company name
    #with no leading or extra spaces"""
    #a = "COMPANY NAME:"
    #if text.find(a) > -1:
        #name = text[len(a):len(text)].strip()
    #else:
        #name = text.strip()
    #return name

def sic_code(text):
    """Receives raw data and finds the SIC code in AICPA files
    Need to check in which row it is, since some files have an address"""
    a = "SIC_CODE:"
    max_elen = len("SIC CODE: 737; 7374")+1
    #check = len("717 RIDGEDALE AVENUE; EAST HANOVER, NJ 07936")
    com_sep = [":", ";"]
    #print(text[0])
    #print(len(text[0]))
    if len(text[0]) <= max_elen:
        if text[0].find(a) > -1:
            code = text[0][len(a):len(text[0])].strip()
        else:
            code = text[0].strip()
    else:
        if text[0].find(a) > -1:
            code = text[0][len(a):len(text[0])].strip()
        else:
            code = text[0].strip()

    print(code)

def if_find(value, text):
    """Takes a string and looks for a value
    if found it returns the string without that value
    else it returns the stripped string"""
    if text.find(value) > -1:
        return text[len(value):len(text)].strip()
    else:
        return text.strip()


def write_file(data):
    """Writes all data to file"""
    #If no GVKEY or document contains many files, then one CSV file per document
    #For files containing less


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
                a = [file_path_open]
                a.extend(fsttotal(file_path_open, os.path.splitext(i)[0]))
                a.append(getText(file_path_open, a))
                #print(a[4])
                #print(*a[4][0], sep='\n')
                #print(len(a[4]))
                #print(a[2])
                #print(a)
                #print(i for i in a[4])
        else:
            a = [file_path_a]
            a.extend(fsttotal(file_path_a, os.path.splitext(file)[0]))
            a.append(getText(file_path_a, a))
            #print(*a[4][0], sep='\n')
            #print(len(a[4]))
            #print(a[2])
            #print(a)
    return a


def fipath(gvkey, path):
    """Function delivers path to files to open"""
    path = os.path.abspath(path)
    if os.path.isdir(path) == False:
        file_name = os.path.splitext(os.path.basename(path))[0]
        # get file name without path or extension
        a = [path]
        a.extend(fsttotal(path, file_name))
        a.append(getText(path, a))
        print(*a[4][0], sep='\n')
        parseText(a[2], a[4])
        #print(len(a[4]))
        #print(a[2])
        #print(a)
        return a
    else:
        return file_loop(path)
