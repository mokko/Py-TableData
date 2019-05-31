'''
Load a table and search it for a needle. List the cells (x,y) which contain the string. 
 search.py --input in.xls --search needle
'''

import argparse
import os
import sys
from TableData import TableData

def error (msg):
    print ('Error: ', msg,"\n")
    sys.exit (1)

parser = argparse.ArgumentParser()
parser.add_argument('-i', '--input', required=True)
parser.add_argument('-s', '--search', required=True)

args = parser.parse_args()

if not os.path.isfile(args.input):
     error ('Input file not found')

td=TableData.load_table(args.input)
td.show()
res=td.search(args.search)
print (res)

