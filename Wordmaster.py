""" Script to read AICPA Word Files"""
# Parse word files and separate them into individual documents
# Save individual files as txt for data gathering
import os

path_1 = "C:/Users/Panqiao/Documents/Research/AICPA/Files to separate/NO GVKEY"
print(os.listdir(path_1))
lpath_l = os.listdir(path_1)
for i in lpath_l:
    print(i)

