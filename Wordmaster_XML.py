import docx2txt
import zipfile
import re
import os
import xml.etree.ElementTree as ET
nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
file_experiment = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\Annual_Reports_-_Corporate_(AICPA)__1972-1982011-05-07_23-05.docx'
file_experiment = os.path.abspath(file_experiment)

#text = docx2txt.process(file_experiment)
#print(text)


document = zipfile.ZipFile(file_experiment)
filelist = document.namelist()
#for i in filelist:
    #print(i)


def qn(tag):
    """
    Stands for 'qualified name', a utility function to turn a namespace
    prefixed tag name into a Clark-notation qualified tag name for lxml. For
    example, ``qn('p:cSld')`` returns ``'{http://schemas.../main}cSld'``.
    Source: https://github.com/python-openxml/python-docx/
    """
    prefix, tagroot = tag.split(':')
    uri = nsmap[prefix]
    return '{{{}}}{}'.format(uri, tagroot)


def xml2text(xml):
    """
    A string representing the textual content of this run, with content
    child elements like ``<w:tab/>`` translated to their Python
    equivalent.
    Adapted from: https://github.com/python-openxml/python-docx/
    """
    text = u''
    root = ET.fromstring(xml)
    for child in root.iter():
        print(child)
        if child.tag == qn('w:t'):
            t_text = child.text
            text += t_text if t_text is not None else ''
        elif child.tag == qn('w:tab'):
            text += '\t'
        elif child.tag in (qn('w:br'), qn('w:cr')):
            text += '\n'
        elif child.tag == qn("w:p"):
            text += '\n\n'
    print(child.tag)
    return text

text = u''
header_xmls = 'word/header[0-9]*.xml'
for fname in filelist:
    if re.match(header_xmls, fname):
        print(fname)
        print(xml2text(document.read(fname)))
        text += xml2text(document.read(fname))

#print(text)

#import zipfile
#try:
    #from xml.etree.cElementTree import XML
#except ImportError:
    #from xml.etree.ElementTree import XML


file_experiment = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\Annual_Reports_-_Corporate_(AICPA)__1972-1982011-05-07_23-05.docx'
file_experiment = os.path.abspath(file_experiment)

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
TABLE = WORD_NAMESPACE + 'tbl'
ROW = WORD_NAMESPACE + 'tr'
CELL = WORD_NAMESPACE + 'tc'
HEADER = WORD_NAMESPACE + 'headerReference'
TOP = WORD_NAMESPACE + 'top'
SECT = WORD_NAMESPACE + 'sectPr'
HEAD = WORD_NAMESPACE + 'hdr'


def get_docx_text(path):
    document = zipfile.ZipFile(path)
    for i in dir(document):
        print(i)
    print(document.filelist)
    for i in document.filelist:
        print(i)
    xml_content = document.read('word/document.xml')
    xml_header = document.read('word/header9.xml')
    print(xml_header)
    document.close()
    tree = XML(xml_content)
    tree1 = XML(xml_header)
    for i in dir(tree.iter(PARA)):
        print(i)
    for i in tree.iter(SECT):
        print(i.text)
        for e in i.iter():
            print(e)
    paragraphs = []
    for i in tree1.iter(TEXT):
        print(i.text)
    for i in tree.iter(PARA):
        for e in i.iter(SECT):
            print(e.attrib)
    for paragraph in tree.getiterator(PARA):
        texts = [node.text
                 for node in paragraph.getiterator(TEXT)
                 if node.text]
        if texts:
            paragraphs.append(''.join(texts))

    return '\n\n'.join(paragraphs)

#for i in getiterator.get_docx_text(file_experiment):
    #print(i)
#get_docx_text(file_experiment)
#print(get_docx_text(file_experiment))