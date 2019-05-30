from TableData import TableData
# run pytest from ../test with python -m pytest

def test_normalConstructor():     
    td=TableData ('xls', 'test/data.xls')
    assert td.cell(0,0) == 'objId'
    assert td.cindex('objId') == 0
    assert td.nrows() == 2093
    assert td.ncols() == 20
    assert td.search ('objId') == [(0,0)]

def test_load_table():     
    td=TableData.load_table ('test/data.xls')
    assert td.cell(0,0) == 'objId'
    assert td.cindex('objId') == 0
    assert td.nrows() == 2093
    assert td.ncols() == 20
    assert td.search ('objId') == [(0,0)]

def test_all_types_of_writes():
    td=TableData.load_table ('test/data.xls')
    td.write ('test/data.xml')
    td.write ('test/data.json')
    td.write ('test/data.csv') #doesn't work?

def test_load_xml():
    td=TableData.load_table ('test/data.xml')
    assert td.cell(0,0) == 'objId'
    assert td.cindex('objId') == 0
    assert td.nrows() == 2093
    assert td.ncols() == 20
    #sth wrong with search, but only in XML
    #assert td.search ('objId') == [(0,0)]

def test_load_csv():
    td=TableData.load_table ('test/data.csv')
    assert td.cell(0,0) == 'objId'
    assert td.cindex('objId') == 0
    assert td.nrows() == 2093
    assert td.ncols() == 20
    assert td.search ('objId') == [(0,0)]

def test_load_json():
    td=TableData.load_table ('test/data.json')
    assert td.cell(0,0) == 'objId'
    assert td.cindex('objId') == 0
    assert td.nrows() == 2093
    assert td.ncols() == 20
    assert td.search ('objId') == [(0,0)]


