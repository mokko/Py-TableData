from TableData import TableData

##
## Transformations that are not properly abstracted
##


class Ffexport (TableData):

    def erwerbNotizAusgabe (self):
        if not td.colExists('ErwerbNotizAusgabe'):
            eNotiz=td.addCol ('ErwerbNotizAusgabe')
            
        inst=td.cindex('VerwaltendeInstitution')
        eDatum=td.cindex('ErwerbDatum') # exception if col doesnt exist, but can be empty
        eArt=td.cindex('Erwerbungsart')
        
        for rid in range(1, td.nrows()):
            Inst=td.cell (inst,rid)
            EDatum=td.cell(eDatum, rid)
            EArt=td.cell(eArt, rid)
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
            td.table[rid][eNotiz]=text

    #variation over default
    def rewrite_credits(self):
        cid=td.cindex('Credits')
        vi=td.cindex('VerwaltendeInstitution')
        for rid in range(1, td.nrows()):
            if not td.cell (cid,rid):
                self.table[rid][cid]=self.table[rid][vi]


if __name__ == '__main__':
    td=Ffexport.load_table ('data/WAF55 Gestalter XSL 20190529.xls', 'v')
    print (type(td))
    exit (0);
    td.erwerbNotizAusgabe()
    td.delCol('ErwerbDatum')
    td.delCol('Erwerbungsart')
    td.clean_whitespace ('OnlineBeschreibung')
    td.delCol('IdentNrSort')
    td.rewrite_credits()
    
    td.write('data.xml')
    td.write('data.csv')
