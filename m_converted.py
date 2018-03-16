"""Find converted files and move them to a separate folder for backup"""
#Find files that could not be converted and move them to a separate folder
import glob
import os
path = input('Path to files to move:\n') #ask for path to apply the script
#file_type = input('Type:\n') #ask type of file to move
#path_move = input('Path to move files to:\n')
docx_ext = r'**/*.docx'
doc_ext = r'**/*.doc'
#y = os.path.join(path,docx_ext)
#y = os.path.join(path,docx_ext)
#y1 = os.path.join(path,file_type_to)


path_docx = glob.glob(os.path.join(path,docx_ext), recursive = True) #compile path of files to .dox files

def check_path():
    """Check if any files failed to convert"""
    #If it has been converted, then erase from path
    #Save path of files not converted
    ask = input('Do you want to check if files have failed to convert before? Write yes or no\n')
    if ask == 'yes':
        b = glob.glob(os.path.join(path, doc_ext), recursive=True)  # compile path of files to .doc files
        #aa = [os.path.abspath(i) for i in a] #do not need this...
        #b = [os.path.abspath(i) for i in b]
        path_s = [os.path.splitext(i)[0] for i in b] #split file name from extension
        path_checks = [os.path.splitext(i)[0] for i in path_docx]
        path_failed = [x for x in path_s if x not in path_checks] #check for filenames in path_s not in path_checks
        path_converted = [x for x in path_s if x in path_checks] #check for filenames in path_s not in path_checks
        out_file1 = os.path.abspath("out{}.doc".format(path_failed))
        out_file2 = os.path.abspath("out{}.docx".format(path_converted))
    else:
        print("Did not search for missing files")
        return None
    return out_file1, out_file2

paths = check_path()
print(paths[0])
print(paths[1])

#print(path_failed)
#for i in path_failed:
   # print(i)