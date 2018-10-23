'''
Created on 11/10/2011

@author: pantonio
'''

import csv
#import sys
import logging
from ui.get_two_files import App
import tkinter.tix
from tkinter import ttk
import os
import stat

GUI_TITLE="Preprocesos EERR"

START_ROW='start_row'
ACCT_PREFIX='account_prefix'
OUTPUT_IDX='output'
INPUT_IDX='input'
HEADERS='headers'
ARR_MONTH=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']

#Data list
DATA_DATE='date'
DATA_COMPTE='compte'
DATA_TYPE='type'
DATA_COMMENT='comment'
DATA_AREA='area'
DATA_COST_CENTER='cost_center'
DATA_ITEM='item'
DATA_EFF_DATE='eff_date'
DATA_ANALISYS_DATE='analisys_date'
DATA_REFERENCE='reference'
DATA_REF_DATE='ref_date'
DATA_EXP_DATE='exp_date'
DATA_DEBIT='debit'
DATA_CREDIT='credit'
DATA_BALANCE='balance'
DATA_BRANCH='branch'
DATA_ACCT_NUM = 'acct_num'
DATA_ACCT_DESC = 'acct_desc'
DATA_ITEM_DESC = 'item_desc'
DATA_MONTH='month'
DATA_STATUS='status'
DATA_COMPANY='company'
DATA_DESC_AREA='desc_area'
DATA_BRAND='brand'
DATA_DET_EERR='det_eerr'
DATA_EERR='eerr'

class LibroMayor(App):
    def __init__(self, master, fd):
        self.params_fd = fd
        self.items = None
        self.init_params = None
        self.areas = None
        self.eerr = None
        self.libro_mayor()
        App.__init__(self, master,self.proc_fn)
        
    def proc_fn(self):
        self.load_libro_mayor(self.get_file1(), self.get_file2())

    def libro_mayor(self):
        #fd=open(file, 'r')
        #line 1 items file
        s = self.params_fd.readline().replace('\n','')
        self.items = load_item_file(s)
        #line 2 libro mayor init file
        s = self.params_fd.readline().replace('\n','')
        self.init_params = load_libro_mayor_format(s)
        #line 3 area file
        s = self.params_fd.readline().replace('\n','')
        self.areas = load_area_file(s)
        #line 4 estado resultado file
        s = self.params_fd.readline().replace('\n','')
        self.eerr = load_eerr_file(s)
        #line 5 root dir
        #self.rootdir = self.params_fd.readline().replace('\n','')

    def load_libro_mayor(self, in_file, out_file):
        f_stat = os.stat(in_file)
        pb = ttk.Progressbar(self.frame, length = f_stat[stat.ST_SIZE])
        pb.pack()
        
        fd=open(in_file,'r', newline='')
        fd2=open(out_file,'w', newline='')
        row_count=1
        reader = csv.reader(fd, delimiter=';')
        writer = csv.writer(fd2, delimiter=';', dialect='excel')
        row = next(reader)
        #empresa
        company=row[0].strip()
        status="REAL"
        current_acct=""
        curr_acct_desc=""
        output_fmt = self.init_params[OUTPUT_IDX]
        input_idx = self.init_params[INPUT_IDX]
        header = self.init_params[HEADERS]
        line = list()
        for i in range(len(output_fmt)):
            line.append(header[output_fmt[i+1]])
        writer.writerow(line)
    
        for row in reader:
            first = row[0].strip()
            if (first.find(self.init_params[ACCT_PREFIX])==0):
                current_acct = first[len(ACCT_PREFIX)+1:].strip()
                desc_pos = current_acct.find(' ')
                curr_acct_desc = current_acct[desc_pos+1:].strip()
                current_acct = current_acct[:desc_pos]
                if current_acct  in self.eerr:
                    current_eerr_det=self.eerr[current_acct]
                elif current_acct[:6] in self.eerr:
                    current_eerr_det=self.eerr[current_acct[:6]]
                else:
                    current_eerr_det='N/A'
            elif len(first)> 0 and current_acct != "":
                line=list() 
                if len(row) < 16:
                    line.append("ERROR en linea " + row_count.__str__() + " faltan datos: ")
                    logging.error(line)
                    logging.error(str)
                else: 
                    for i in range(len(output_fmt)):
                        if i in output_fmt:
                            if output_fmt[i] == DATA_ACCT_NUM:
                                line.append(current_acct)
                            elif output_fmt[i] == DATA_ACCT_DESC:
                                line.append(curr_acct_desc)
                            elif output_fmt[i] == DATA_ITEM_DESC:
                                if row[input_idx[DATA_ITEM]]in self.items:
                                    line.append(self.items[row[input_idx[DATA_ITEM]]])
                                else:
                                    line.append('N/A')
                            elif output_fmt[i] == DATA_MONTH:
                                date = row[input_idx[DATA_DATE]].strip()
                                #line+='N/A'
                                sm=date.find('/')+1
                                if sm != -1:
                                    em=date.find('/', sm)
                                    if em != -1:
                                        line.append(ARR_MONTH[int(date[sm:em])-1])
                            elif output_fmt[i] == DATA_COMPANY:
                                line.append(company)
                            elif output_fmt[i] == DATA_STATUS:
                                line.append(status)
                            elif output_fmt[i] == DATA_DESC_AREA:
                                area_cod=row[input_idx[DATA_AREA]]
                                if area_cod in self.areas:
                                    line.append(self.areas[area_cod][1])
                                else:
                                    line.append('N/A')
                            elif output_fmt[i] == DATA_BRAND:
                                area_cod=row[input_idx[DATA_AREA]]
                                if area_cod in self.areas:
                                    line.append(self.areas[area_cod][0])
                                else:
                                    line.append('N/A')
                            elif output_fmt[i] == DATA_DET_EERR:
                                line.append(current_eerr_det)
                            elif output_fmt[i] == DATA_EERR:
                                line.append('')
                            else:
                                line.append(row[input_idx[output_fmt[i]]])
                writer.writerow(line)
            row_count += 1
        fd.close()
        fd2.close()
        logging.info("Process done. " + row_count.__str__()+" rows processed.")
    
def load_libro_mayor_format(file):
    logging.debug("Formatting options")
    fd=open(file,'r')
    s = fd.readline()
    params=dict()
    output_fmt=dict()
    input_idx=dict()
    header = dict()
    params[OUTPUT_IDX]=output_fmt
    params[INPUT_IDX]=input_idx
    params[HEADERS]=header
    while(len(s)> 0):
        s=s.replace('\n','')
        if len(s)>0:
            if s[0]!= '#':
                eq_pos = s.find('=')
                param_cd = s[:eq_pos].strip()
                param_val = s[eq_pos+1:].strip()
                if param_cd[:14]=="output_format_":
                    output_fmt[int(param_cd[14:])]=param_val
                elif param_cd[:6] == "input_":
                    input_idx[param_cd[6:]]=int(param_val)-1
                elif param_cd[:5] == "head_":
                    header[param_cd[5:]]=param_val.strip()
                else:
                    params[param_cd]=param_val
        s = fd.readline()
    fd.close()
    return params

def load_item_file(file):
    logging.debug("Loading item file")
    fd=open(file, 'r', newline='')
    reader = csv.reader(fd,delimiter=';')
    next(reader)
    item_cod_desc = dict()
    for row in reader:
        if len(row) == 2:
            item_cod_desc[row[0]] = row[1]
            #logging.debug(row)
    fd.close()
    return item_cod_desc
'''
Load a file which map an area code with a brand
and a company group
'''
def load_area_file(file):
    logging.debug("Loading Area file")
    fd=open(file, 'r', newline='')
    reader = csv.reader(fd, delimiter=';')
    next(reader)
    area_cod_desc = dict()
    for row in reader:
        area_cod_desc[row[0]]=[row[1],row[2]]
    fd.close()
    return area_cod_desc

'''
Load a file which map a profit and loss report code prefix 
with its description  
'''
def load_eerr_file(file):
    logging.debug("Loading EERR file")
    fd=open(file, 'r', newline='')
    reader = csv.reader(fd, delimiter=';')
    next(reader)
    eerr_cod_desc = dict()
    for row in reader:
        eerr_cod_desc[row[0]]=row[1]
    fd.close()
    return eerr_cod_desc

'''
def libro_mayor(fd):
    #fd=open(file, 'r')
    #line 1 items file
    s = fd.readline().replace('\n','')
    items = load_item_file(s)
    #line 2 libro mayor init file
    s = fd.readline().replace('\n','')
    init_params = load_libro_mayor_format(s)
    #line 3 area file
    s = fd.readline().replace('\n','')
    areas = load_area_file(s)
    #line 4 estado resultado file
    s = fd.readline().replace('\n','')
    eerr = load_eerr_file(s)
    #line 5 libro mayor file
    s = fd.readline().replace('\n','')
    while 1:
        files = get_files('Libro mayor')
        if files == None:
            break
        input_file = files[0]
        output_file = files[1]
        if input_file is None:
            print("No selecciono ningun archivo") #. Desea continuar?"
        elif output_file is None:
            print("Indique nombre de archivo de salida.")
            #output_file= filesavebox(msg, GUI_TITLE, default, filetypes)
            #print(output_file)
        else:
            load_libro_mayor(items, init_params, areas, eerr, input_file, output_file)
            sys.exit(0)
    '''
def start(fd):
    root = tkinter.tix.Tk(className='Libro mayor')
    LibroMayor(root, fd)
    root.mainloop()

    
