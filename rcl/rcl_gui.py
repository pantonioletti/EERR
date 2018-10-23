'''
Created on 14/11/2011

@author: pantonio
'''
from easygui import fileopenbox
#import sys

def get_file_name():        
    default='C:\dev\projects\ScriptsUtil\data\*.csv'
    msg='Seleccione archivo a procesar'
    title='Procesando archivos para EERR'
    filetypes=["*.csv",]
    file = fileopenbox(msg,title,default,filetypes)
    print(file)