'''
Convert table to output format.

    tdconv.py --input infile --out out.xml

Conversion to format indicated by file extension. Extension are converted to lowercase.
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
parser.add_argument('-o', '--out', required=True)
args = parser.parse_args()
args.out=args.out.lower()

if not os.path.isfile(args.input):
     error ('Input file not found')

td=TableData.load_table(args.input, 'verbose')
td.write(args.out)

