""" Script to read AICPA Word Files
Not meant to read other types of files
Will have to create other functions for that"""
# First gather identifying data and place it into a spreadsheet
import os
import docx
#import string
import glob


def getText(filename, file_details):
    """Function gathers text fom docx file"""
    #Will add a variable that takes the list of paragraph numbers within the file
    #And looks for the needed text to gather data
    doc = docx.Document(filename)
    fullText = []
    para = doc.paragraphs
    #print("Have entered GETEXT")
    for i in file_details[3]:
        newText = []
        for j in range(i, i+29):
           (newText.append(para[j].text) if para[j].text != '' else None)
        fullText.append(newText)
    return fullText


def parseText(num_docs, text):
    """Parse the text gotten from geText"""
    id_data = []
    months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
    months.extend([s.strip()+'.' for s in months])
    #months_2 = [s.strip()+'.' for s in months]
    print(months)
    for i in text:
        print(i)
        #print(set(i))
        name = if_find("COMPANY NAME:", i[1])
        sic = sic_code(i[2:4])
        date = find_strings_lists(i, months)
        print(date)
        #date = if_find(":", date, option = 1)
        #print(date)
        if date == None:
            input("press enter:")
        #if len(name) >= 64: #special case do not forget
            #print(name)
            #for j in i:
                #print(j)

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
    a = "SIC CODE:"
    com_sep = str.maketrans("","",";: ") #check how to make this better
    try:
        int(if_find(a, text[0]).translate(com_sep))
        code = if_find(a, text[0])
    except:
        code = if_find(a, text[1])
    try:
        int(code.translate(com_sep))
    except:
        code = "NA"
        print("FIX SIC CODE")
    return code

def if_find(value, text, option = 0):
    """Takes a string and looks for a value
    if found it returns the string without that value (and before)
    else it returns the stripped string"""
    if option == 0:
        if text.find(value) > -1:
            return text[len(value):len(text)].strip()
        else:
            return text.strip()
    if option == 1:
        if text.find(value) > -1:
            return text[text.find(value)+1:len(text)]
        else:
            return text.strip()


def find_strings_lists(text, terms, terms1 = ""):
    c1 = next((s for s in text for s1 in terms if s1 in s.split()), None)
    print(c1)
    c1 = if_find(":", c1, option = 1)
    return c1
        #print(i.find(any(terms)))
#mylist = ['abc123', 'def456', 'ghi789']
#sub = 'abc'
#next((s for s in mylist if sub in s), None) # returns 'abc123'




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
                parseText(a[2], a[4])
        else:
            a = [file_path_a]
            a.extend(fsttotal(file_path_a, os.path.splitext(file)[0]))
            a.append(getText(file_path_a, a))
            parseText(a[2], a[4])
            #if parseText(a[2], a[4]) == 'error':
               #break
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
        parseText(a[2], a[4])
        return a
    else:
        return file_loop(path)
