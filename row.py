#!/usr/bin/python

import TableData
import shutil
import os

td=TableData('xls','mume.xls')
c1=td.table.cindex('MMPfad')
c2=td.table.cindex('MMDateiname')
c3=td.table.cindex('MMErweiterung')

for r in td.table
	fpath=os.path.join(r[c1], r[c2]) + '.' + r[c3]
    if os.path.isfile(fpath):
        verbose ('FOUND ', fpath)
        #todo: dryrun if args.x:
            cpFile(fpath)
    else:
        print ('NOT found ', fpath)

td.writeXml('out.xml')






def cpFile (filepath):
    if not os.path.isfile(filepath):
        verbose ('not found %', filepath)
        return False
    try:
        shutil.copy2(filepath, '.') # copy2 attempts to preserve file info; why not

    except:    
        print("Unexpected error:", sys.exc_info()[0])