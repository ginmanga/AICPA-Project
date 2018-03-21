""" Script to read AICPA Word Files"""
# First gather identifying data and place it into a spreadsheet
import os
import docx
#from docx import Document
file_experiment = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\Annual_Reports_-_Corporate_(AICPA)__1972-1982011-05-07_23-05.docx'
file_experiment = os.path.abspath(file_experiment)
file_test = docx.Document(file_experiment)
sections = file_test.sections

#print(file_test.paragraphs[7].text == '')
#print(file_test.paragraphs[7].text.isspace())
print(len(sections))
#for i in dir(file_test.sections.count):
    #print(i)
#print("HHHHHHHHHHHEEERR")
#print(file_test.paragraphs[3])

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
        break
    count += 1
#for i in file_test.sections:
    #if i.text != '' and i.text.isspace() == False:
        #print(i.text)
        #print(count)
        #break
    #count += 1


#import zipfile
#try:
    #from xml.etree.cElementTree import XML
#except ImportError:
    #from xml.etree.ElementTree import XML


#file_experiment = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\Annual_Reports_-_Corporate_(AICPA)__1972-1982011-05-07_23-05.docx'
#file_experiment = os.path.abspath(file_experiment)

#WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
#WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
#PARA = WORD_NAMESPACE + 'p'
#TEXT = WORD_NAMESPACE + 't'
#TABLE = WORD_NAMESPACE + 'tbl'
#ROW = WORD_NAMESPACE + 'tr'
#CELL = WORD_NAMESPACE + 'tc'
#HEADER = WORD_NAMESPACE + 'headerReference'
#TOP = WORD_NAMESPACE + 'top'
#SECT = WORD_NAMESPACE + 'sectPr'
#HEAD = WORD_NAMESPACE + 'hdr'


#def get_docx_text(path):
    #document = zipfile.ZipFile(path)
    #for i in dir(document):
        #print(i)
    #print(document.filelist)
    #for i in document.filelist:
        #print(i)
    #xml_content = document.read('word/document.xml')
    #xml_header = document.read('word/header9.xml')
    #print(xml_header)
    #document.close()
    #tree = XML(xml_content)
    #tree1 = XML(xml_header)
    #for i in dir(tree.iter(PARA)):
        #print(i)
    #for i in tree.iter(SECT):
        #print(i.text)
        #for e in i.iter():
            #print(e)
    #paragraphs = []
    #for i in tree1.iter(TEXT):
        #print(i.text)
    #for i in tree.iter(PARA):
        #for e in i.iter(SECT):
            #print(e.attrib)
    #for paragraph in tree.getiterator(PARA):
        #texts = [node.text
                 #for node in paragraph.getiterator(TEXT)
                 #if node.text]
        #if texts:
            #paragraphs.append(''.join(texts))

    #return '\n\n'.join(paragraphs)

#for i in getiterator.get_docx_text(file_experiment):
    #print(i)
#get_docx_text(file_experiment)
#print(get_docx_text(file_experiment)) """