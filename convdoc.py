"""Script to Convert old word files into .docx files"""
import glob
import win32com.client
import os

word = win32com.client.Dispatch("Word.Application")
word.visible = 0
#def ask_path(x):
x = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\NO GVKEY'
#x.replace('\','\\')
#print(os.path.abspath(x))
#x = os.path.abspath(x)+r'\*.doc'
#print(x)
#path1 = glob.glob(x, recursive=False)
#y='C:\\Users\\Panqiao\\Documents\\Research\\AICPA\\Files to separate\\NO GVKEY\\*.doc'
#print(y)

def usr_path():
    path = input('Path to files to convert:\n')
    return path

#print(usr_path())
pathn = usr_path()
y= pathn+'\*.doc'
print(y)
#path = glob.glob('C:\\Users\\Panqiao\\Documents\\Research\\AICPA\\Files to separate\\NO GVKEY\\*.doc', recursive=False)
path = glob.glob(y, recursive = True)
print(path)
for i in path:
    print(i)
for i in path:
    break
    in_file = os.path.abspath(i)
    #print(in_file)
    wb = word.Documents.Open(in_file)
    #print(os.path.splitext(in_file)[0])
    fn = os.path.splitext(in_file)[0] # takes name of file without extension
    fn_docx = check + ".docx" # adds docx extension
    #print(check2)
    #out_file = os.path.abspath("out{}.docx".format(check))
    out_file = check2
    #print(out_file)
    wb.SaveAs2(out_file, FileFormat=16) # file format for docx
    wb.Close()

word.Quit()