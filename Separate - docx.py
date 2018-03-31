"""Gather Text and Data from AICPA files using Wordmaster as input"""
from Wordmaster_nextgen import fipath

#directory = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\Annual_Reports_-_Corporate_(AICPA)__1972-1982011-05-07_23-05.docx'
#directory = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\GVKEY'
#directory = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\NO GVKEY'
#directory = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\Annual_Reports_-_Corporate_(AICPA)__1972-1982011-05-07_23-05.docx'


directory = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\GVKEY'
directory = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\NO GVKEY'
directory = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\Annual_Reports_-_Corporate_(AICPA)__1972-1982011-05-07_23-05.docx'
#print(fipath(0, directory))

directory = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\GVKEY'
directory = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\14848199310kad.docx'

#print(fipath(0, directory))
fipath(0, directory)

dir = [r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\GVKEY', r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\NO GVKEY']
#for i in dir:
    #fipath(0, i)