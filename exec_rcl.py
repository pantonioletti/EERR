'''
Created on 14/10/2011

@author: pantonio
'''
from rcl.libro_mayor import start
import logging
import sys
import os
import os.path as op

params = sys.argv
log_level=logging.WARNING
for p in params:
    if p[:4]== 'log=':
        if p[4:5] == 'D':
            log_level = logging.DEBUG
        elif p[4:5] == 'I':
            log_level = logging.INFO
        elif p[4:5] == 'E':
            log_level = logging.ERROR
        elif p[4:5] == 'C':
            log_level = logging.CRITICAL
    elif p[:3] == 'wd=':
        wd = p[3:]
        if op.exists(wd):
            os.chdir(wd)

logging.basicConfig(filename='rcl_eerr.log',format='%(asctime)s %(levelname)s: %(message)s',level=log_level)
try:
    fd=open('params.txt', mode='r')
except IOError:
    print('File params.txt not found')
start(fd)
fd.close()