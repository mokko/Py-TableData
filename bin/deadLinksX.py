# This Python file uses the following encoding: utf-8
'''
Expect .xls file with mume links, check if files exists and report on missing files in EXCEL
'''

import datetime
import os
import argparse
from TableData import TableData

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input', required=True)
    #parser.add_argument('-o', '--output', required=False)
    #write output to STDOUT in shell fashion
    args = parser.parse_args()
    
    if not os.path.isfile(args.input):
        error ('Input file not found')
        
    td=TableData('xls',args.input)
    c1=td.cindex('multimediaPfadangabe') #  MMPfadangabe
    c2=td.cindex('multimediaDateiname') # MMDateiname
    c3=td.cindex('multimediaErweiterung') # MMErweiterung
    mulId=td.cindex('mulId') # MMMulId
    objId=td.cindex('objId') # ObjId
    
    notFound=0
    incomplete=0
    noPath=0
    xls_line=6
    del td.table[0] # headers

    from xlrd import open_workbook    
    from xlutils.copy import copy 
    from xlwt import Workbook
    rb = open_workbook(args.input,formatting_info=True)
    wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)
    report=wb.add_sheet('ToteLinks')

    
    for r in td.table:
        fpath=os.path.join(r[c1],r[c2] ) + '.' + r[c3]

        #Mume-DS m�ssen keine Pfade haben. Das ist ok, Wir suchen hier nur F#lle, wo Pfade eingegeben wurden
        #Wenn ein Pfad eingegeben ist, dann muss er auch vollständig sein  
        if r[c1] or r[c2] or r[c3]: 
            if r[c1]=='' or r[c2]=='' or r[c3]=='':
                report.write(xls_line,2,r[mulId])
                report.write(xls_line,3,fpath)
                report.write(xls_line,4,'unvollständiger Pfad')
                xls_line += 1
                incomplete += 1
            else:
                if not os.path.isfile(fpath):
                    notFound += 1 
                    report.write(xls_line,2,r[mulId])
                    report.write(xls_line,3,fpath)
                    report.write(xls_line,4,'toter Link')
                    xls_line += 1
        else: #no path filled in whatsoever
            noPath+=1
   
    #y,x?
    report.write(0,0,str(datetime.datetime.now()))
    report.write(1,0, 'Zeilen (OK)')
    report.write(1,1, td.nrows())
    report.write(2,0, 'kein Pfad (OK)')
    report.write(2,1, noPath)
    report.write(3,0, 'Unvollständiger Pfad (Fehler)')
    report.write(3,1, incomplete)
    report.write(4,0, 'tote Links (Fehler)')
    report.write(4,1, notFound)
    report.write(5,2, 'mulId')
    report.write(5,3, 'Pfad')
    report.write(5,4, 'Diagnose')
    
    #wb.save(args.input + '.out' + os.path.splitext(args.input)[-1]) 
    wb.save(args.input) 
