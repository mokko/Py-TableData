#!/usr/bin/python
import sys # not necessary...
import os 
import argparse
import shutil
from TableData import TableData

'''
Copy multimedia resources to new directory. Resources are renamed according to objID.
    cpMume.py --input mume.xsl # copies to directory of input file
    
    Expects the following excel columns: MMMulId, Standardbild, MMPfadangabe, MMDateiname, MMErweiterung
    
    TODO: --output param
'''
verbose = 1
   
def verbose (msg):
    if verbose: 
        print (msg)
        
def cpFile (in_p, out_p):
    if os.path.isfile(fpath):
        verbose ('FOUND '+ in_p)
        try: 
            shutil.copy2(in_p, out_p) # copy2 attempts to preserve file info; why not
        except:
            print("Unexpected error:", sys.exc_info()[0])
    else:
            verbose ('NOT found '+ in_p)
        
if __name__ == "__main__":
    
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input', required=True)
    parser.add_argument('-o', '--output', required=False)

    args = parser.parse_args()
    if not os.path.isfile(args.input):
        error ('Input file not found')

    outpath = os.path.dirname(args.input)
    #print ("outdir:"+outpath)
    if args.output:
        outpath=args.output
    
    if not os.path.isdir (outpath):
        print ("Error: Output dir not found")
#
#
#
    td=TableData('xls',args.input)
    c1=td.cindex('MMPfadangabe')
    c2=td.cindex('MMDateiname')
    c3=td.cindex('MMErweiterung')
    mulId=td.cindex('MMMulId')
    objId=td.cindex('ObjId')
    Stdbild=td.cindex('Standardbild')
    
    c=0
    for r in td.table:
        if r[mulId] == r[Stdbild]:
            fpath=os.path.join(r[c1],r[c2] ) + '.' + r[c3]
            out=os.path.join(outpath, str(r[objId])+ '.' + r[c3])
            cpFile(fpath, out)
            #verbose ('-->'+ out)
            c+=1
    verbose ('Tried to copy %i files' % c)
exit (0)

