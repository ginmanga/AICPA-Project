"""Script to Convert old word files into .docx files"""
import glob
import win32com.client
import os

word = win32com.client.Dispatch("Word.Application")
word.visible = 0
#x = r'C:\Users\Panqiao\Documents\Research\AICPA\Files to separate\GVKEY'
#x.replace('\','\\')
#print(os.path.abspath(x))
#x = os.path.abspath(x)+r'\*.doc'
#print(x)
#path1 = glob.glob(x, recursive=False)


path = input('Path to files to convert:\n')
file_type = input('Type to convert:\n')
file_type_to = input('Type to convert to:\n')

y = os.path.join(path,file_type)
y1 = os.path.join(path,file_type_to)
#path = glob.glob('C:\\Users\\Panqiao\\Documents\\Research\\AICPA\\Files to separate\\GVKEY\\**/*.doc', recursive=False)
#s = r'C:\\Users\\Panqiao\\Documents\\Research\\AICPA\\Files to separate\\GVKEY\\**/*.doc'
#print(s)
#**/*.docx
path = glob.glob(y, recursive = True)
path_check = glob.glob(y1, recursive = True)

#print(path)
#print(path_check)
def check_path(a,b):
    """Check file to check if converted"""
    #If it has been converted, then erase from path
    #print(a)
    ask = input('Do you want to check if files have been converted before? Write yes or no\n')
    if ask == 'yes':
        #path_s = [s.strip('.DOC') for s in a]
        #path_checks = [s.strip('.docx') for s in b]
        aa = [os.path.abspath(i) for i in a]
        b = [os.path.abspath(i) for i in b]
        path_s = [os.path.splitext(i)[0] for i in aa]
        path_checks = [os.path.splitext(i)[0] for i in b]
        path = [x for x in path_s if x not in path_checks]
    else:
        return a
    #print(path_s)
    #print(path_checks)
    #print(path)
    return path
newpath = check_path(path,path_check)
print(newpath)
path = newpath
#for i in newpath:
    #print(i)

for i in path:
    in_file = os.path.abspath(i)
    print(in_file)
    try:
        wb = word.Documents.Open(in_file)
    except:
        print("Could not open")
        print(in_file)
        continue

    fn = os.path.splitext(in_file)[0] # takes name of file without extension
    fn_docx = fn + ".docx" # adds docx extension
    #print(check2)
    #out_file = os.path.abspath("out{}.docx".format(check))
    out_file = fn_docx
    #print(out_file)
    wb.SaveAs2(out_file, FileFormat=16) # file format for docx
    wb.Close()

word.Quit()