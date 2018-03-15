"""Find converted files and move them to a separate folder for backup"""
import glob
import os
path = input('Path to files to move:\n') #ask for path to apply the script
#file_type = input('Type:\n') #ask type of file to move
path_move = input('Path to move files to:\n')
docx_ext = r'**/*.docx'
y = os.path.join(path,docx_ext)

#y = os.path.join(path,docx_ext)
#y1 = os.path.join(path,file_type_to)


path = glob.glob(y, recursive = True) #compile path of files to .dox files

def check_path(a):
    """Check if any files failed to convert"""
    #If it has been converted, then erase from path
    #print(a)
    ask = input('Do you want to check if files have failed to convert before? Write yes or no\n')
    if ask == 'yes':
        doc_ext = r'**/*.doc'
        b = glob.glob(os.path.join(path, doc_ext), recursive=True)  # compile path of files to .doc files
        aa = [os.path.abspath(i) for i in a]
        b = [os.path.abspath(i) for i in b]
        path_s = [os.path.splitext(i)[0] for i in aa]
        path_checks = [os.path.splitext(i)[0] for i in b]
        path = [x for x in path_s if x not in path_checks]
    else:
        return a
    return path