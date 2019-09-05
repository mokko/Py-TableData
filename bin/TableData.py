import os

'''
TableData deals with data that comes from MS Excel, csv, xml. More precisely, it expects
a single table which has headings in the first row. It converts between these formats and usually keeps 
information on a round trip between those formats identical.

TableData also allows for simple transformations, like dropping a column.


CONVENTIONS
*cid is column no or column id
*rid is row no or row id
*cell refers the content of a cell, a cell is represented by cid|rid, as two integers or (not sure yet) a tuple or a list 
*cname is the column name (in row 0)

NOTE
* Column names have to be unique
* (x|y) not rows x cols 
* Currently internal cells do have a type, which may be flattened to str if output is type agnostic.
* cid and rid begins with 0, so first cell is 0|0, but ncols and nrows start at 1. Strangely enough, sometimes that is convenient.
* interface prefers cname over cid

LIMITATIONS
Data is stored in memory (in a two dimensional list of lists), so max. size depends on available memory (ram). 

WHAT NOT TO DO
I will NOT allow conversion INTO Excel xsl format, only reading from it. 

I will not abstract this thing too far. I write it for my current Excel version and the csv flavor that I
need (e.g. csv is escaped only for values that contain commas). I don't need multiple Excel sheets, 
formatting in Excel, lots of types in Excel.

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
        ''' Raise exception if column names (cnames) are not unique. '''
        if len(set(self.table[0])) != len(self.table[0]):
            raise Exception('Column names not unique')

    def __init__ (self, ingester, infile, verbose=None):
        self._verbose=verbose
        self.table=[] # will hold sheet in memory as list of list
        if ingester == 'mpx':
            self.MPXParser(infile)
        elif ingester == 'xml':
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
        ''' File extension aware ingester

            td=TableData.load_table(path)
        
        This is an alternative constructor to _init_.'''    
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



    '''
    Parses mpx into two-dimensional list where 
    a) only one record type (sammlungsobject) 
    b) every record (distinct id) has one row
    c) group elements become multiple entries in a single cell 
    d) parameters are flattened in the usual way (elem/@param becomes elemParam)
    '''
    def MPXParser (self,infile):
        self.verbose ('xml infile %s' % infile)
        import xml.etree.ElementTree as ET
        tree = ET.parse(infile)
        tags=set() # distinct list for columns
        data=[] # list of dictionaries to temporarily store the data
        
        #What happens with repeated attributes (Wiederholfelder) in this format? 
        #Most get deleted/overwritten! Only the last one survives 
        
        #turn records into rows, and aspects into into columns
        #we also turn aspect@attributes to aspectAttributes columns
        for elem in tree.iter('{http://www.mpx.org/mpx}sammlungsobjekt'):
            #print (elem.tag, elem.attrib)
            record={}# single record
            for each in 'objId', 'exportdatum':
                record[each]=elem.attrib[each]
                tags.add(each)
            
            for aspect in elem.findall('*'):
                aspectNoNS=aspect.tag.split("}")[1]   
                #print ('DICT'+aspectNoNS+':'+str(aspect.text))
                tags.add(aspectNoNS)
                if aspectNoNS in record:
                    record[aspectNoNS]=self._appender (record[aspectNoNS],aspect.text)
                else:
                    record[aspectNoNS]=aspect.text

                for param in aspect.attrib:
                    paramNotation=aspectNoNS+param[0].upper()+param[1:]
                    #print ('!DICT'+paramNotation+':'+str(aspect.attrib.get(param)))
                    tags.add(paramNotation)
                    if paramNotation in record:
                        record[paramNotation]=self._appender(record[paramNotation], aspect.attrib.get(param))
                    else:
                        record[paramNotation]=aspect.attrib.get(param)
                    '''            
                    Problem: When we rewrite group elements (Wiederholfelder) that have uneven number of attributes, we lose information 
                    since it is no longer clear to which entry the attributes refer. 
                    
                    Group elements are rewritten as single elements with multiple entries by listing them with a separator, e.g.
                        <geoBezug>Indien</geoBezug> 
                        <geoBezug>Assam</geoBezug>
                    becomes
                        <geoBezug>Indien; Assam<geoBezug>
                    (Of course the separator implies some masking issues which we don't need to deal with at the moment.)
                    
                    There is a problem with uneven number of attributes (in our source format empty parameters are left out):
                        <geoBezug bezeichnung="Land">Indien</geoBezug> 
                        <geoBezug>Assam</geoBezug>
                    becomes
                        <geoBezug>Indien; Assam<geoBezug>
                        <geoBezugBezeichnung>Land</geoBezugBezeichnung>
        
                    It is no longer clear what entry in geoBezug Land refers to. I want to fix that by marking empty entries:
                        <geoBezug>Indien; Assam<geoBezug>
                        <geoBezugBezeichnung>Land;</geoBezugBezeichnung>
                    
                    So first we need to determine the attribute group, e.g. the attributes that belong together. In this case the attribute 
                    group belonging to the mother element geoBezug. It consists of geoBezugBezeichnung and geoBezugBemerkung. (We only have 
                    to do that for group elements and only if they have attributes, not for every element.)
                    
                    For now let's say we check for attribute group every time we write an attribute. We could check the tags set since it has 
                    already been updated. Alternatively, we could also examine the xml doc using xpath etc. But it's not enough to check only 
                    the attributes of same element. We also have to check all the attributes of the sister elements of the attribute group:
                    e.g. all sibling geoBezug elements. 
                    
                    So perhaps it's easier to check the set of flattened attributes. We know the geoBezug is the mother element, so we just
                    go thru every tag known so far and check if it begins with substring geoBezug. The list that results describes the
                    group attributes (with or without the mother element). Now we compare the number of entries for both and add empty attributes 
                    until they have the same number. The next time we write an attribute we do the same check, so we don't need to know what's
                    in the future.  
                    '''
                    groupAttributes=[tag for tag in tags if aspectNoNS in tag]
                    print ("||"+aspectNoNS+'/'+str(groupAttributes))
                    soll=0 # no of semicolons (i.e. entries-1)
                    #analyze status quo
                    for each in groupAttributes:
                        if each in record:
                            count=record[each].count(';') # count semicolons
                            if count > soll:
                                soll=count
                        else:
                            record[each]='' #make null params in group explicit
                    #add necessary semicolons        
                    for each in groupAttributes:
                        count=record[each].count(';') # count semicolons
                        if count < soll:
                            record[each]=record[each]+(';'*(soll-count))
                        print ("  %s:%s (S/I:%s/%s)" % (each,record[each], soll, count ))

            #print (record)
            data.append(record) 
        #print (sorted(tags))

        self.table.append(sorted(tags)) #add columns

        for record in data:
            col=[]
            for tag in self.table[0]:
                if tag in record:
                    col.append(record[tag])
                    #print (tag+':'+str(record[tag]))
                else:
                    col.append('')
            self.table.append(col)
        #self.verbose (self.table)


    def _appender (self, strA, strB):
        return strA.rstrip()+'; '+strB.rstrip()

    def JSONParser (self, infile):
        import json
        self.verbose ('json infile %s' % infile)
        json_data = open(infile, 'r').read()
        self.table = json.loads(json_data)
        
##
## read table data, but NO manipulations
##
    def ncols(self):
        ''' Returns integer with number of columns in table data
            ncols=self.ncols()
        '''
        return len(self.table[0])
    
    def nrows (self):
        '''  Returns integer with number of rows in table data 
            nrows=self.nrows()
        '''
        return len(self.table)
    
    def cell (self, col,row):
        ''' For a given columnn and row, return the corresponding cell:
            cell=td.cell(col,row)

        Throws exception if col or row are not integer or out of range.
        What happens on empty cell? Returns '' not none.
        
        I stick to x|y format, although row|col might be more pythonic.'''
        
        try:
            return self.table[row][col]
        except:
            self.verbose ('%i|%i doesnt exist' % (col, row))
            exit (1)

    def cindex (self,needle):
        ''' Returns the cid for the column name 'needle'.
        
        Throws 'not in list' if 'needle' is not a column name (cname). '''
        
        return self.table[0].index(needle)

    def colExists (self, cname):
        '''Returns True if cname is the name of an existing column name, False if it doesn't.'''
        try:
            self.table[0].index(cname)
            return True
        except:
            return False


    def search (self, needle): 
        ''' Returns list of tuples (cid,rid) that contain the needle.
        r=td.search(needle) '''

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
        ''' Returns list/set of rows that contain the needle for the given col.
            td.search(cname, needle)
        
        UNTESTED
        '''
        results=()
        c=cindex(cname)
        for rid in range(0, self.nrows()): 
            if needle in self.cell(c,rid):
                results.append(rid)
        return results        

    def show (self, cname=None):
        ''' show table or column
            self.show()      # print representation of table
            self.show(cname) # print column
        
        Really print? Why not. It's meant for quick debugging.'''

        if cname is not None:
            cid=self.cindex(cname)
        
        for row in self.table:
            if cname is None: 
                print (row)
            else:
                print ('|'+str(row[cid])+'|')
                    
        print ('Table size is %i x %i (cols x rows)' % (self.ncols(), self.nrows()))            

##
## SIMPLE UNCONDITIONAL TRANSFORMATIONS 
## 

    def delRow (self, rid):
        ''' Drop a row by number. (Rid is zero-based)''' 
        self.table.pop(rid)
        #print ('row %i deleted' % rid)

    def delCol (self, cname):  
        ''' Drop a column by cname '''
        
        c=self.cindex (cname)    
        for r in range(0, self.nrows()):
            self.table[r].pop(c)


    def addCol (self,name):
        ''' Add a new column called at the end of the row with given name. 

        cid=self.addCol('newName')

        Respective cells will be empty. Returns the cid of the new column, 
        same as cindex(cname). '''
        #update 
        self.table[0].append(name) 
        self._uniqueColumns()
        for rid in range(1, self.nrows()):
            self.table[rid].append('') # append empty cells for all rows
        return len(self.table[0])-1 # len starts counting at 1, but I want 0

    def clean_whitespace (self,cname):
        '''Remove certain whitespace (windows style new lines and double spaces)'''
        cid=self.cindex(cname)
        for rid in range(1, self.nrows()):
            self.table[rid][cid]=self.table[rid][cid].replace('\r\n', ' ').replace('  ', ' ')


    def set_col (self, cname, value):
        '''Write value in column with cname. Every row will have same content. Existing content 
        will be overwritten'''
        cid=self.cindex(cname)
        for rid in range(1, self.nrows()):
            self.table[rid][cid]=value
       
##
##  MORE COMPLEX MANIPULATION
##

    def _sortOrder (r): pass

    def sortByCol (self, cname):
         cid=self.cindex(cname)
         #self.table.sort(key=_sortOrder)
         self.table.sort(reverse=True)
    
    def setCellAIfColBContains (self,cnameA, cnameB, needle, cell): 
        ''' Write cell in colA if col B contains the needle.
            self.delCellAIfColBContains (A,B, 'bla') '''
        

    def delCellAIfColBContains (self,cnameA, cnameB, needle):
        ''' In each row, empty cell in column A if column B contains the needle.
            self.delCellAIfColBContains (A,B, 'bla')
        UNTESTED
        '''
        colA=self.cindex(cnameA)
        colB=self.cindex(cnameB) 
        for rid in range(1, self.nrows()):
            if self.table[rid][colB] == needle:
                self.verbose ('del %s if Col %s contains needle %s' % (cnameA, cnameB, needle))
                selt.table[rid][colA]=''

    #def delCellAIfColBEQ (self,cnameA, cnameB, needle): pass

    def delRowIfColContains (self, cname, needle): 
        '''
        Delete row (not just cell) if column equals the value 'needle'

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
    def renameCol (self, cnameOld, cnameNew):
        ''' Renames column cnameOld into cnameNew
            cid=self.renameCol (old, new)
        Returns column index of the column.'''
        cid=self.cindex(cnameOld)
        self.table[0][cid]=cnameNew
        return cid    

    def default_per_col (cname, default_value):
        '''
        Default Value: if cell is empty replace with default value
            self.default_per_col ('status', 'filled')
        '''
        cid=td.cindex(cname)
        for rid in range(1, td.nrows()):
            if not td.cell (cid,rid):
                self.table[rid][cid]=default_value
            

###
### converting to outside world
###
    
    def _outTest(self,out):
        if os.path.exists(out):
            self.verbose('Output exists already, will be overwritten: %s' %out)
    

    def write (self, out):
        ''' Write table to file, format is picked according to extension
            self.write(outfile) #writes csv, xml or json '''
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
        ''' Writes table to file in csv format:
            self.writeCSV (outfile)
        
        Values with commas are quoted. Output is UTF-8 
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
        Writes table to file in xml format
            self.writeXML(outfile)
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
        
        #don't need cnames here, so start at 1, but then write all columns in first row 
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
        ''' Writes table to json file. 
            self.writeJSON(outfile)
        JSON doesn't have date type, hence default=str'''    
        import json
        self._outTest(out)

        f = open(out, 'w')
        with f as outfile:
            json.dump(self.table, outfile, default=str)
        self.verbose ('json written to %s' % out)
        
        
if __name__ == '__main__': 
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--input', required=True)
    #parser.add_argument('-o', '--output', required=False)
    args = parser.parse_args()
        
    td=TableData.load_table(args.input, 'v')
    #td.show()
    pre, ext = os.path.splitext(args.input)

    #td.write(pre + '.xml')
    td.write(pre + '.csv')
    #td.write(args.output)
 