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

CONVENTIONS
*cid is column no or column id
*rid is row no or row id
*cell is represented by cid|rid, as two integers or (not sure yet) a tuple or a list 
*cname is the column name (in row 0)


NOTE
* (x|y) not rows x cols 
* Currently internal cells do have a type
* cid begins with 0, so first cell is 0|0, but ncols and nrows start at 1   
* nrows and ncols count in Python fashion, beginning with 0
* interface prefers cname over cid

WHAT NOT TO DO
I will NOT allow conversion INTO Excel xsl format, only reading from it. 

UNICODE
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

The first row will have all columns, even empty ones. The other rows usually omit empty elements with empty values.

'''

class TableData:
    def verbose (self, msg):
        if self._verbose: 
            print (msg)
    
    
    def _uniqueColumns (self):
        '''
        raise exception if column names (cnames) are not unique
        '''
        if len(set(self.table[0])) != len(self.table[0]):
            raise Exception('Column names not unique')

    def __init__ (self, ingester, infile, verbose=None):
        self._verbose=verbose
        if ingester == 'xml':
            self.XMLParser(infile)
        elif ingester == 'xls':
            self.XLRDParser(infile)
        elif ingester == 'csv':
            self.CSVParser(infile)
        elif ingester == 'json':
            self.JSONParser(infile)
        #todo: modern excel
        else:
            raise Exception ('Ingester %s not found' % ingester)
        self._uniqueColumns()
    
#
# INGESTERS (xml, csv)
#
    
    def load_table (path, verbose=None):
        '''
        File extension aware ingester

            td=TableData.load_table(path)
        
        This is an alternative to _init_. Is this pythonic enough? 
        '''    
        ext=os.path.splitext(path)[1][1:]
    
        return TableData (ext, path,verbose)
    
    
    def XLRDParser (self, infile):
        '''
        Parses old excel file into tableData object. Only first sheet.

        Dont use this directly, use 
            td=TableData('xsl', infile)
            td=TableData.load=table(infile)
        instead
        
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
        self.table=[] # will hold sheet in memory as list of list
        self.verbose ('xlrd infile %s' % infile)


        #if not os.path.isfile(infile):
        #    raise Exception ('Input file not found')    
        wb = xlrd.open_workbook(filename=infile, on_demand=True)
        sheet= wb.sheet_by_index(0)
        
        #I'm assuming here that first row consist only of text cells?
        
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
                #    self.verbose ("%i/%i contains a comma" % (c,r) )   
                row.append(val)
            self.table.append(row)
        wb.unload_sheet(0) #unload xlrd sheet to save memory   


    def CSVParser (self,infile): 
        import csv
        self.table=[] # will hold sheet in memory as list of list
        self.verbose ('csvParser: ' + str(infile))
        with open(infile, mode='r', newline='') as csvfile:
            incsv = csv.reader(csvfile, dialect='excel')
            for row in incsv:
                self.table.append(row)
                #self.verbose (str(row))
        
   
    def XMLParser (self,infile):
        #It is practically impossible to reconstruct the full list of columns from xml file
        #if xmlWriter leaves out empty elements. Instead, I write them at least for first row.
        self.table=[] # will hold sheet in memory as list of list; overwrite
        self.verbose ('xml infile %s' % infile)
        import xml.etree.ElementTree as ET
        tree = ET.parse(infile)
        for row in tree.iter("row"):
            c=0
            cnames=[]
            col=[]
            for e in row.iter():
                if e.tag !='row':
                    #self.verbose ('%s %s' % (e.tag, e.text))
                    if len(self.table) == 0:
                        #need to create 2 rows from first row in xml
                        cnames.append(e.tag)
                    col.append(e.text)
            if len(self.table)  == 0:        
                self.table.append(cnames)
            self.table.append(col)
        #self.verbose (self.table)

    def JSONParser (self, infile):
        self.table=[] # will hold sheet in memory as list of list; overwrite
        import json
        self.verbose ('json infile %s' % infile)
        json_data = open(infile, 'r').read()
        self.table = json.loads(json_data)
        
##
## read table data, but manipulations
##
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
            td.cell(col,row) # returns [0,1]

        Throws exception if col or row are not integer or out of range.
        
        I stick to x X y format, although (row X col) could also make sense
        '''
        return self.table[row][col]       

    def cindex (self,needle):
        '''
        Returns the column index (c) for column name 'needle'.
        
        Throws 'not in list' if 'needle' is not a column name (cname).
        '''
        return self.table[0].index(needle)

    def search (self, needle): 
        '''
        Returns list of cells [cid,rid] that contain the needle.
            r=td.search(needle) # (1,1)
        
        
        tuples, lists? I am not quite sure!    
        '''
        results=[]
        for rid in range(0, self.nrows()): 
            for cid in range(0, self.ncols()):
                cell=self.cell(cid, rid)
                #self.verbose ('ce:'+str(cell))
                if str(needle) in str(cell):
                    #self.verbose ("%i/%i:%s->%s" % (cid, rid, cell, needle))
                    results.append ((cid,rid))
        return results

    def search_col (self, cname, needle): 
        '''
        Returns list/set of rows that contain the needle for the given col.
            td.search(cname, needle)
        '''
        results=()
        c=cindex(cname)
        for rid in range(0, self.nrows()): 
            if needle in self.cell(c,rid):
                results.append(rid)


    def show (self):
        '''
        print representation of table
        
        Really print? Why not.
        '''
        for row in self.table:
                print (row)
                
        print ('Table size is %i x %i (cols x rows)' % (self.ncols(), self.nrows()))            
##
## Transformations 1 (change table data)
## properly abstracted

    def delRow (self, r):
        '''
        Drop a row by number.
        
        Need to remake the index to cover the hole.
        ''' 
        #r always means rid
        self.table.pop(r)
        #print ('row %i deleted' % r)

    def delCol (self, cname):  
        '''
        Drop a column by cname
        
        (Not tested.)
        '''
        
        c=self.cindex (cname)    
        for r in range(0, self.nrows()):
            self.table[r].pop(c)
                    
    def delCellAIfColBEq (self,cnameA, cnameB, needle):
        '''
        empty cell in column cnameA if value in column cnameB equals needle in every row
        
        untested
        '''
        colA=self.cindex(cnameA)
        colB=self.cindex(cnameB) 
        for r in range(1, self.nrows()):
            if self.table[r][colB] == needle:
                self.verbose ('delCellAifColBEq A:%s, B:%s, needle %s' % (cnameA, cnameB, needle))
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
                enotiz="Erworben  %edatum durch %eart"
                enotiz="Erworben am %edatum durch %eart"
            
        #whole columns could be deleted afterwards
        #but not as part of this rule
        #self.delCol(Erwerbsdatum)
        #self.delCol(Erwerbsart)
                
###
### converting to outside world
###
    
    def _outTest(self,out):
        if os.path.exists(out):
            self.verbose('Output exists already, will be overwritten: %s' %out)
    

    def write (self, out):
        '''
        write to file with extension-awareness
        '''
        ext=os.path.splitext(out)[1][1:].lower()
        if (ext == 'xml'):
            self.writeXML (out)
        elif (ext == 'csv'):
            self.writeCSV (out)
        elif (ext == 'json'):
            self.writeJSON (out)
        else:
            print ('Format %s not recognized' % ext)    


    def writeCSV (self,outfile):
        '''
        writes data in tableData object to outfile in csv format
        
        Values with commas are quoted. 
        '''
        import csv
        self._outTest(outfile)

        with open(outfile, mode='w', newline='', encoding='utf-8') as csvfile:
            out = csv.writer(csvfile, dialect='excel')
            for r in range(0, self.nrows()):
                row=self.table[r]               
                out.writerow(row)

        self.verbose ('csv written to %s' % outfile)

    
    def writeXML (self,out):
        '''
        writes table data to file out in xml format
        '''
        import xml.etree.ElementTree as ET
        from xml.sax.saxutils import escape
        root = ET.Element("tdx") #table data xml 

        self._outTest(out)

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
        
        #don't need cnames here, so start at 1, but then write columns in first row 
        for r in range(1, self.nrows()):  
            doc = ET.SubElement(root, "row")
            for c in range(0, self.ncols()):      
                cell = self.cell(c,r)
                #print ('x,y: %i/%i: %s->%s ' % (r, c, self.columns[c], cell))
                #for round trip I need empty cells, at least in the first row
                if cell or r == 1:   
                    ET.SubElement(doc, self.table[0][c]).text=escape(str(cell))

        tree = ET.ElementTree(root)
        
        _indent(root)
        tree.write(out, encoding='UTF-8', xml_declaration=True)
        self.verbose ('xml written to %s' % out)


    def writeJSON (self, out):
        '''
        Writes table data in json to file out
        
        JSON doesn't have date type, hence default=str
        '''    
        import json
        self._outTest(out)

        f = open(out, 'w')
        with f as outfile:
            json.dump(self.table, outfile, default=str)
        self.verbose ('json written to %s' % out)
        
        
if __name__ == '__main__':       
        
    td=TableData.load_table ('test/data.xls', 'v')
    td.delCol('ColA')
    td.show()
