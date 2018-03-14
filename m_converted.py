"""Find converted files and move them to a separate folder for backup"""
import glob
import os
path = input('Path to files to convert:\n') #ask for path to apply the script
file_type = input('Type:\n') #ask type of file to move
path_move = input('Path to move files to:\n')
y = os.path.join(path,file_type)
path = glob.glob(y, recursive = True) #compile path of files