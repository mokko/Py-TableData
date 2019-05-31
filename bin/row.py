#!/usr/bin/python

import TableData
import shutil
import os


'''
    Wir wollen Infos aus 2 Spalten in eine neue dritte Spalte schreiben.
    Dafür müssen wir die Spalte erst einmal herstellen.
'''

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





