# This Python file uses the following encoding: utf-8

'''
1. Exportiere alle Standorte nach Excel.
2. Gruppiere alle Standorte pro Objekt-DS. 
3. Finde genau die GD.Standorte, die in Standortgeschichte nicht vorkommen.
4. Liste entsprechende objIds

Notwendige Felder:
GD.AktuellerStandort
GD.St�ndiger Standort
Alle Felder in Standortgeschichte
DatumVon
DatumBis
Standort
StandortDetail
Bearb.Datum
Bearb.Mitarb
Verkn�ofung zu Thesaurus
...

'''

import os
import argparse
from TableData import TableData

if __name__ == "__main__":
    
    parser.add_argument('-i', '--input', required=True)
    args = parser.parse_args()
    
    if not os.path.isfile(args.input):
        error ('Input file not found')


    td=TableData('xls',args.input)
    
    objId_no=td.cindex('objID')
    GD_AktSto_no=td.cindex('GD.AktuellerStandort')
    GD_StSto_no=td.cindex('GD.StändigerStandort')
    Sto=td.cindex('Standort')

    #all distinct objId exactly once
    objIds=set()
    for r in td.table:
        objIds.add (r[objId_no])

    #find exactly the DS with a specific objId
    noEqualInSTOHistory=set() #should be a list    
    for objId in objIds:
        hits=0
        for r in td.table:
            if r[objId_no] == objId:
                if GD_AktSto is Null: #only the first time
                    GD_AktSto=r[GD_AktSto_no]
                    GD_StSto=r[GD_StSto_no]

                if r[Sto]==GD_AktSto or r[Sto]==GD_StSto:
                    hits+=1

        if hits == 0:    
            noEqualInSTOHistory.add(objId)
        
        GD_AktSto=Null # not exactly sure if that works with 

    print (noEqualInSTOHistory)