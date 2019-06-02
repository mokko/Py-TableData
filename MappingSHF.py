from TableData import TableData
'''
Prelimanary mapping for SHF Export

from Classic -> xls -> csv

all logic specific to this mapping should be in this file

methods should be reusable in other mappings
'''

class MappingSHF (TableData):
    
    #def __init__ (self, infile, verbose=None): 
    #    self=TableData.load_table (infile, verbose)
    #def load_file (self, infile,verbose):
    
    #def load_table (path, verbose=None):
            
    
    def erwerbNotizAusgabe (self):
        if not self.colExists('ErwerbNotizAusgabe'):
            eNotiz=self.addCol ('ErwerbNotizAusgabe')
            
        inst=self.cindex('VerwaltendeInstitution')
        eDatum=self.cindex('ErwerbDatum') # exception if col doesnt exist, but can be empty
        eArt=self.cindex('Erwerbungsart')
        
        for rid in range(1, self.nrows()):
            Inst=self.cell (inst,rid)
            EDatum=self.cell(eDatum, rid)
            EArt=self.cell(eArt, rid)
            #print ('EE:'+Inst+EDatum+EArt)
            #mapping data to more sensible format
            if EArt == 'Unbekannt':
                EArt='unbekannte Erwerbungsart'
                
            #Writing German based on available cells
            if len(EDatum) > 4:
                EDatum='am %s' % str(EDatum)
            if Inst and EDatum and EArt:
                text=('Das %s bzw. eine Vorgängerinstitution erwarb das Objekt %s durch %s.' % (Inst,EDatum,EArt))
            elif Inst and EDatum:
                text=('Das %s bzw. eine Vorgängerinstitution erwarb das Objekt %s.' % (Inst,EDatum))
            elif Inst and EArt:
                text=('Das %s bzw. eine Vorgängerinstitution erwarb das Objekt durch %s.' % (Inst,EArt))
            else:
                text=''
            #print ('DD:'+str(rid)+':'+str(eNotiz)+text)
            self.table[rid][eNotiz]=text

    
    def rewrite_credits(self):
        ''' if Credits empty, put Verwaltende Institution in there; do credits point to object information, the object or both?
        ''' 
        cid=self.cindex('Credits')
        vi=self.cindex('VerwaltendeInstitution')
        for rid in range(1, self.nrows()):
            if not self.cell (cid,rid):
                self.table[rid][cid]=self.table[rid][vi]


if __name__ == '__main__':
    map=MappingSHF('csv', 'data.csv', 'v')
    print (type (map))
    #td.erwerbNotizAusgabe()
    #td.delCol('ErwerbDatum')
    #td.delCol('Erwerbungsart')
    map.clean_whitespace ('OnlineBeschreibung')
    map.show(cname='OnlineBeschreibung')
    #td.delCol('IdentNrSort')
    map.rewrite_credits()
    
    #td.write('data.xml')
    #td.write('data.csv')
