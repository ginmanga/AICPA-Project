""" Script to read AICPA Word Files"""
# First gather identifying data and place it into a spreadsheet
import os
import zipfile
try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML


file_experiment = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\Annual_Reports_-_Corporate_(AICPA)__1972-1982011-05-07_23-05.docx'
file_experiment = os.path.abspath(file_experiment)

#WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
TABLE = WORD_NAMESPACE + 'tbl'
ROW = WORD_NAMESPACE + 'tr'
CELL = WORD_NAMESPACE + 'tc'

def get_docx_text(path):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)
    print(tree)

    paragraphs = []
    for paragraph in tree.getiterator(PARA):
        #print(paragraph)
        texts = [node.text
                 for node in paragraph.getiterator(TEXT)
                 if node.text]
        #print(texts)
        #if texts:
            #paragraphs.append(''.join(texts))
    for text in tree.getiterator(TEXT):
        texts = [node.text
                 for node in text.getiterator(TEXT)
                 if node.text]
        #print(texts)
    for table in tree.getiterator(TABLE):
        for row in table.iter(ROW):
            print(''.join(node.text for node in cell.iter(TEXT)))
            #for cell in row.iter(CELL):
                #print(''.join(node.text for node in cell.iter(TEXT)))
        #print(texts)

    return '\n\n'.join(paragraphs)


get_docx_text(file_experiment)
#xml_content = document.read('word/document.xml')
#tree = XML(xml_content)
#print(get_docx_text(file_experiment))
#print(type(get_docx_text(file_experiment)))
#print(dir(get_docx_text(file_experiment)))
#for i in dir(get_docx_text(file_experiment)):
    #print(i)
#for i in tree.get_docx_text(file_experiment):
    #print(i)