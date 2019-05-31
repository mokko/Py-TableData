'''
xpath stuff
'''
#import xml.etree.ElementTree as ET
from lxml import etree

tree = etree.parse('HUF-Standorte Nach Wiederherstellung.xml')
#root = tree.getroot()

#rows = tree.xpath("/tdx/row[Art = 'Aktueller Standort']")
rows = tree.xpath("/tdx/row/objID")

#Group-by in python

#Let's make an index of objIDs
IDs=set()
for r in rows:
#    print (r.text)
    IDs.add (r.text)

print ('Rows: ' + str(len(rows)))
print ('Set: ' + str(len(IDs)))

#keine definitiven aktuellen Standorte --> 3 Stichproben angeguckt, Kein Türkei-Problem.
for i in IDs:
    rows = tree.xpath("/tdx/row[objID = " + i + " and Art='Aktueller Standort' and Status = 'Definitiv']")
    if len(rows) == 0:
        print ('ObjID ', i)

exit(0)
for r in rows:
    for e in r:
        print ('%s:%s' % (e.tag, e.text))


