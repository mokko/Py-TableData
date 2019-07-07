# This Python file uses the following encoding: utf-8
'''

Expect .xls file with mume links, check if files exists and report on missing files

What should the report look like? STDOUT or excel?

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
    
    print ('NICHT GEFUNDEN (mulId:Pfad):')
    notFound=0
    incomplete=0
    noPath=0
    del td.table[0]
    for r in td.table:

        #Wenn ein Pfad eingegeben ist, dann muss er auch vollständig sein  
        if r[c1] or r[c2] or r[c3]: 
            if r[c1]=='' or r[c2]=='' or r[c3]=='':
                print ('%s: unvollständiger Pfad' %r[mulId])
                incomplete += 1
                next
        
        #Mume-DS müssen keine Pfade haben. Das ist ok, Wir suchen hier nur Fälle, wo Pfade eingegeben wurden
        if r[c1]:
            fpath=os.path.join(r[c1],r[c2] ) + '.' + r[c3]
        
            if not os.path.isfile(fpath):
                notFound += 1 
                print ('%s:%s' % (r[mulId], fpath))
        else:
            noPath+=1
    # SUMARY

    now = datetime.datetime.now()
    print ('----------------------')
    print (str(now))
    print ('OK: %i Zeilen' % td.nrows())
    print ('OK: %i kein Pfad' % noPath)
    print ('Fehler: %i unvollständiger Pfad' % incomplete)
    print ('Fehler: %i toter Link' % notFound)
