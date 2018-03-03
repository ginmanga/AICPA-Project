"""Script to Convert old word files into .docx files"""
import glob
import win32com.client
import os

word = win32com.client.Dispatch("Word.Application")
word.visible = 0

path = glob.glob('C:\\Users\\Panqiao\\Documents\\Research\\AICPA\\Files to separate\\NO GVKEY\\*.doc', recursive=False)
#print(path)

for i in path:
    in_file = os.path.abspath(i)
    wb = word.Documents.Open(in_file)
    out_file = os.path.abspath("out{}.docx".format(i))
    wb.SaveAs2(out_file, FileFormat=16) # file format for docx
    wb.Close()

word.Quit()