# -*- coding: utf-8 -*-
"""
Created on Sun May 22 14:12:57 2022

@author: Y-dee
"""

import zipfile
import os
import shutil

# Time taken ~4hrs  

# next time = make in a CLASS
# execute main code in MAIN METHOD of class
# includde try and except statements to troubleshoot sectoins which may throw errors
# Such as reading and writing to files
  
# wordXML extract time spent in document

def zipWordDoc(folderpath, docName, zipName):
    '''
    Parameters
    ----------
    folderpath : string
        Enter folderpath with 'r' preceding the string. Example: r'C:\\Users\\Test
        
    docName: string
        Enter name and .docx file extension of word doc of interest in folderpath

    zipName: string
        Enter name of .docx zipped file that will be converted to zip file using this method
    
    Returns filepath of zipped folder
    -------
    '''
    # makeZipFile in python from wordDoc
    
    zippedDoc = open(zipName, "w")
    zippedDoc.close()
    
    #https://stackoverflow.com/questions/37400974/unicode-error-unicodeescape-codec-cant-decode-bytes-in-position-2-3-trunca
    src =  folderpath + '\\' + docName
    dst = folderpath + str('\\') + zipName
    
    # copy file to destination
    shutil.copyfile(src, dst)
    
    # changed file extension, which then create Windows zipfiles with accessible XMLs
    base = os.path.splitext(dst)[0]
    os.rename(dst, base + '.zip')

    zipFilePath = base + '.zip'
    
    return zipFilePath

def printTotalTimeSpentInWordDocZip(zipFilePath):
    # takes zip folder as input and extracts the total time spent in word document programmatically
    
    zipd = zipFilePath
    test2 = zipfile.ZipFile(zipd)
    
    # https://www.codevscolor.com/list-all-files-zip-python
    #for name in test2.namelist():
    #    print ('%s' % (name))
    
    xmlContent = test2.read('docProps/app.xml')
    
    #parse XML using XML parses/ HTML parser - https://stackoverflow.com/questions/39007743/exact-string-search-in-xml-files
    from bs4 import BeautifulSoup
    bs = BeautifulSoup(xmlContent, "html.parser")
    
    totalTimeTag  = bs.find('totaltime')
        
    timeSpent = str()
    
    for i in str(totalTimeTag):
        if i.isdigit() == True:
            timeSpent = timeSpent+str(i)
    
    timeSpent = int(timeSpent)
    return timeSpent


### MAIN

# why r needed before filepath - https://stackoverflow.com/questions/42654934/need-of-using-r-before-path-name-while-reading-a-csv-file-with-pandas
folderpath = os.getcwd()
docName = "DemonstrationDocument.docx" 
zipName = "zippedFile.docx" 

zipFilePath = zipWordDoc(folderpath, docName, zipName)
    
time = printTotalTimeSpentInWordDocZip(zipFilePath)
print('Time spent in  document: ',time, 'mins')
print('Time spent in document: ',time//60, 'hours ', time%60, 'mins')
