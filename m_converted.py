"""Find converted files and move them to a separate folder for backup"""
#Find files that could not be converted and move them to a separate folder
import glob
import os
#path = input('Path to files to move:\n') #ask for path to apply the script
path = r'C:\Users\Panqiao\Documents\Research\AICPA\fdbgvkey'
#file_type = input('Type:\n') #ask type of file to move
#path_move = input('Path to move files to:\n')
docx_ext = r'**/*.docx'
doc_ext = r'**/*.doc'
docx_extn = r'.docx'
doc_extn = r'.doc'
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
        path_failed = [x+doc_extn for x in path_failed]
        path_converted = [x + doc_extn for x in path_converted]
    else:
        print("Did not search for missing files")
        return None
    return path_failed, path_converted

paths_bad, path_good = check_path()
print(paths_bad)
print(path_good)

def move_good():
    """Call function to move bad files only"""
    #path_bad = input('Type path to move bad files\n')
    path_bad = r'C:\Users\Panqiao\Documents\Research\AICPA\FDBG - DOC FILES'
    for i in path_good:
        tail = os.path.split(i)[1]
        #print(os.path.join(path_bad, tail))

        try:
            os.rename(i, os.path.join(path_bad, tail))
            print("Just moved %s" % (tail))
        except:
            print("Did not move %s" % (tail))



def move_bad():
    """Call function to move bad files only"""
    #path_bad = input('Type path to move bad files\n')
    path_bad = r'C:\Users\Panqiao\Documents\Research\AICPA\FDBG - PROBLEMS'
    for i in paths_bad:
        tail = os.path.split(i)[1]
        #print(os.path.join(path_bad, tail))
        os.rename(i, os.path.join(path_bad, tail))


#move_bad()
move_good()