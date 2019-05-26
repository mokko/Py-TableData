import os

'''
TableData deals with data that comes from MS Excel, csv, xml. More precisely, it expects
a single table which has headings in the first row. It converts between these formats and usually keeps 
information on a round trip between those formats identical.

TableData also allows for simple transformations, like dropping a column.

Data is stored in memory (in a two dimensional list of lists), so there are size limitations. The max. size 
depends on available memory (ram). 

I will not abstract this thing too far. I write it for my current Excel version and the csv flavor that I
need (e.g. csv is escaped only for values that contain commas). I don't need multiple Excel sheets, 
formatting in Excel, lots of types in Excel.

I will NOT allow conversion INTO Excel xsl format, only reading from it. 

I am going for UTF-8 encoding, but not sure I have it everywhere yet. xlrd is internally in UTF16LE, I believe.

Roundtrip Exceptions
*date

XML Format made by TableData is
<tdx>
   <row>
       <cnameA>cell value</cnameA>
       <cnameB>cell value</cnameB>
       ...
    </row>
</tdx>

The first row will have all columns, even empty ones. The other rows may omit empty elements with empty values.

'''


verbose=1
def verbose (msg):
    if verbose: 
        print (msg)

class TableData:
    def _uniqueColumns (self):
        '''
        raise exception if column names (cnames) are not unique
        '''
        if len(set(self.table[0])) != len(self.table[0]):
            raise Exception('Column names not unique')

    def __init__ (self, ingester, infile):
        if ingester == 'xml':
            self.xmlParser(infile)
        elif ingester == 'xls':
            self.xlrdParser(infile)
        elif ingester == 'csv':
            self.csvParser(infile)
        #todo: modern excel
        else:
            raise Exception ('Ingester %s not found' % ingester)
    
#
# INGESTERS (xml, csv)
#
    
    
    def xlrdParser (self, infile):
        '''
        Parses old excel file into tableData object. Only first sheet.
        
        xlrd uses UTF16. What comes out of here?
        
        TO DO: 
        1. better tests for
        -Unicode issues not tested
        -Excel data fields change appearance
        2. conversion/transformation stuff
        '''
        
        import xlrd
        import xlrd.sheet
        from xlrd.sheet import ctype_text
        verbose ('xlrd infile %s' % infile)


        self.table=[] # will hold sheet in memory as list of list
        #if not os.path.isfile(infile):
        #    raise Exception ('Input file not found')    
        wb = xlrd.open_workbook(filename=infile, on_demand=True)
        sheet= wb.sheet_by_index(0)
        
        #self.ncols=sheet.ncols #Excel starts counting at 1
        #self.nrows=sheet.nrows #Excel starts counting at 1
        
        #I'm assuming here that first row consist only of text cells
        
        #start at r=0 because we want to preserve the columns
        for r in range(0, sheet.nrows): #no
        
            row=[]
            for c in range(sheet.ncols):
        
                cell = sheet.cell(r, c) 
                cellTypeStr = ctype_text.get(cell.ctype, 'unknown type')
                val=cell.value

                #convert cell types -> dates look changed, but may not be (seconds since epoch)!
                if cellTypeStr == "number":
                    val=int(float(val))
                elif cellTypeStr == "xldate":
                    val=xlrd.xldate.xldate_as_datetime(val, 0)
                #Warn if comma -> to check if escaped correctly -> quoting works        
                #if ',' in str(val):
                #    verbose ("%i/%i contains a comma" % (c,r) )   
                row.append(val)
            self.table.append(row)
        wb.unload_sheet(0) #unload xlrd sheet to save memory   
        self._uniqueColumns()


    def csvParser (self,infile): 
        import csv
        verbose ('csvParser: ' + str(infile))
        with open(infile, mode='r', newline='') as csvfile:
            incsv = csv.reader(csvfile, dialect='excel')
            self.table=[]
            for row in incsv:
                self.table.append(row)
                #verbose (str(row))
        self._uniqueColumns()
        
   
    def xmlParser (self,infile):
        #It is practically impossible to reconstruct the full list of columns from xml file
        #if xmlWriter leaves out empty elements. Instead, I write them at least for first row.
        verbose ('xml infile %s' % infile)
        self.table=[] # will hold sheet in memory as list of list
        import xml.etree.ElementTree as ET
        tree = ET.parse(infile)
        for row in tree.iter("row"):
            c=0
            cnames=[]
            col=[]
            for e in row.iter():
                if e.tag !='row':
                    verbose ('%s %s' % (e.tag, e.text))
                    if len(self.table) == 0:
                        #need to create 2 rows from first row in xml
                        cnames.append(e.tag)
                    col.append(e.text)
            if len(self.table)  == 0:        
                self.table.append(cnames)
            self.table.append(col)
        self._uniqueColumns()
        #verbose (self.table)
        
#
# read table data
#
    def ncols(self):
        '''
        Returns integer with number of columns in table data
        '''
        return len(self.table[0])
    
    def nrows (self):
        '''
        Returns integer with number of rows in table data
        '''
        return len(self.table)
    
    def cell (self, col,row):
        '''
        Return a cell for col,row.

        Throws exception if col or row are not integer or out of range.
        '''
        
        return self.table[col][row]       

    def cindex (self,needle):
        '''
        Returns the column index (c) for column name 'needle'.
        
        Throws 'not in list' if 'needle' is not a column name (cname).
        '''
        return self.table[0].index(needle)

    def search (self, cname, needle): pass

##
## Transformations 1 (change table data)
## properly abstracted

    def delRow (self, r):
        '''
        Drop a row by row no (which begins at 0)
        ''' 
        #r always means rid
        self.table.pop(r)
        #print ('row %i deleted' % r)

    def delCol (self, cname):  
        '''
        Drop a column by cname
        
        (Not tested.)
        '''
        
        for r in range(0, self.nrows()):
            c=self.cindex (cname)    
            self.table[r].pop[c]
                    
    def delCellAIfColBEq (self,cnameA, cnameB, needle):
        '''
        empty cell in column cnameA if value in column cnameB equals needle in every row
        
        untested
        '''
        colA=self.cindex(cnameA)
        colB=self.cindex(cnameB) 
        for r in range(1, self.nrows()):
            if self.table[r][colB] == needle:
                verbose ('delCellAifColBEq A:%s, B:%s, needle %s' % (cnameA, cnameB, needle))
                selt.table[r][colA]=''

    def delCellAIfColBContains (self,col_a, col_b, needle): pass

    def delRowIfColContains (self, cname, needle): 
        '''
        Delete row if column equals the value 'needle'

        Should we use cname or c (colId)?
        '''
        #cant loop thru rows and delete one during the loop     
        col=self.cindex(cname)

        #it appears that excel and xlrd start with 1
        #todo: not sure why I have shave off one here!
        r=self.nrows()-1 
        while r> 1:
            #print ('AA%i/%i: ' % (r,col))
            cell=self.cell (r, col)
            if needle in str(cell):
                #print ('DD:%i/%s:%s' % (r, cname, cell))
                #print ('delRowIfColEq: needle %s found in row %i'% (needle, r))
                self.delRow(r)
            r -=1    
        
           
    def delRowIfColEq (self,col, needle): pass

    
    def addCol (self,name):
       '''
       Add a new column called name at the end of the row. 
       Cells with be empty.
       
       Untested
       '''
       #update 
       self.table[0].append(name) 

    def renameCol (self, cnameOld, cnameNew):
        '''
        renames column cnameOld into cnameNew
        '''
        c=self.cindex(cnameOld)
        self.table[0][c]=cnameNew    
##
## Transformations that are not properly abstracted --> callbacks
##
    def RewriteErwerbNotiz (self):
        '''
        Practice run for a ErnerbNotiz@Ausgabe produced by me.

        Untested
        '''
        for r in range(1, self.nrows()):
            #Ich definiere mal eine Form
            #im Augenblick kein Ver�u�erer, nur Datum/Jahr und EArt
            #Erworben durch %Erwerbsart %Erwerbsdatum  
            edatum=self.cname('Erwerbsdatum') 
            eart=self.cname('Erwerbsart') 
            enotiz=self.cname('Erwerbnotiz')
            equali==self.cname('ErwerbQualifikator')

            edatum=self.table[r][edatum]
            eart=self.table[r][eart]
            enotiz=self.table[r][enotiz]
            equali=self.table[r][equali]

            #Do nothing if equali=='Ausgabe'
            if equali != 'Ausgabe':
                enotiz="Erworben durch %edatum %eart"
                enotiz="Erworben am %edatum durch %eart"
            
        #whole columns could be deleted afterwards
        #but not as part of this rule
        #self.delCol(Erwerbsdatum)
        #self.delCol(Erwerbsart)
                
###
### converting to outside world
###
    
    def writeCsv (self,outfile):
        '''
        writes data in tableData object to outfile in csv format
        
        Values with commas are quoted. 
        '''
        import csv
        with open('out.csv', mode='w', newline='', encoding='utf-8') as csvfile:
            out = csv.writer(csvfile, dialect='excel')
            for r in range(0, self.nrows()):
                row=self.table[r]               
#               row=[]
#               for c in range(self.ncols):
#                   row.append(self.cell(r, c)) 
                out.writerow(row)

        verbose ('csv written to %s' % outfile)

    
    def writeXml (self,outfile):
        '''
        writes data in tableData object to outfile in xml format
        '''
        import xml.etree.ElementTree as ET
        from xml.sax.saxutils import escape
        root = ET.Element("tdx") #table data xml 

        def _indent(elem, level=0):
            i = "\n" + level*"  "
            if len(elem):
                if not elem.text or not elem.text.strip():
                    elem.text = i + "  "
                if not elem.tail or not elem.tail.strip():
                    elem.tail = i
                for elem in elem:
                    _indent(elem, level+1)
                if not elem.tail or not elem.tail.strip():
                    elem.tail = i
            else:
                if level and (not elem.tail or not elem.tail.strip()):
                    elem.tail = i
        
        #doc = ET.SubElement(root, "multimediaObjekt")
        #don't need cnames here, so start at 1
        for r in range(1, self.nrows()):  
            doc = ET.SubElement(root, "row")
            for c in range(0, self.ncols()):      
                cell = self.cell(r, c)
                #print ('x,y: %i/%i: %s->%s ' % (r, c, self.columns[c], cell))
                #for round trip I need empty cells, at least in the first row
                if cell or r == 1:   
                    ET.SubElement(doc, self.table[0][c]).text=escape(str(cell))

        tree = ET.ElementTree(root)
        
        _indent(root)
        tree.write(outfile, encoding='UTF-8', xml_declaration=True)
        verbose ('xml written to %s' % outfile)


#
# really simplistic testing
#
if __name__ == '__main__':
    td=TableData('xls','HUFO-Schau-Südsee-EM-mume-Export.xls')
    td.writeCsv('out.csv')
    td.writeXml('out.xml')
    td=TableData('csv', 'out.csv')
    td=TableData('xml', 'out.xml')
    
    verbose ("----------")
    exit()

    td=TableData('xls','HUFO-Schau-Südsee-EM-mume-Export.xls')
    print ('ncols: %s' % td.ncols())
    print ('nrows: %s' % td.nrows())
    print ('cell 0,1: %s' % (td.cell (0,1)))
    #print (td.cell ('bla',1))# clear error msg
    #print (td.cell (100,10000)) # clear error msg

    print (td.table[0])
    print ('cid for mulId: %s' % td.cindex('mulId')) 
    #print ('cid for mulId: %s' % td.cindex('mulIdss')) 
    #td.writeXml('out.xml')
    td.delRowIfColEq('multimediaUrhebFotograf', 'Peter Jacob')
    td.writeCsv('out.csv')