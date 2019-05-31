from TableData import TableData
# run pytest from ../test with python -m pytest

def test_normalConstructor():     
    td=TableData ('xls', 'test/data.xls')
    assert td.cell(0,0) == 'ColA'
    assert td.cindex('ColB') == 1
    assert td.nrows() == 3
    assert td.ncols() == 2
    assert td.search ('ein') == [(0,1)]

def test_load_table():     
    td=TableData.load_table ('test/data.xls')
    assert td.cell(0,0) == 'ColA'
    assert td.cindex('ColB') == 1
    assert td.nrows() == 3
    assert td.ncols() == 2
    assert td.search ('ein') == [(0,1)]

def test_all_types_of_writes():
    td=TableData.load_table ('test/data.xls')
    td.write ('test/data.xml')
    td.write ('test/data.json')
    td.write ('test/data.csv') 
    td.write ('test/data.json') 

def test_load_xml():
    td=TableData.load_table ('test/data.xml')
    assert td.cell(0,0) == 'ColA'
    assert td.cindex('ColB') == 1
    assert td.nrows() == 3 # xml has only two rows, but internal representation has three
    assert td.ncols() == 2
    assert td.search ('ein') == [(0,1)]

def test_load_csv():
    td=TableData.load_table ('test/data.csv')
    assert td.cell(0,0) == 'ColA'
    assert td.cindex('ColB') == 1
    assert td.nrows() == 3
    assert td.ncols() == 2
    assert td.search ('ein') == [(0,1)]

def test_load_json():
    td=TableData.load_table ('test/data.json')
    assert td.cell(0,0) == 'ColA'
    assert td.cindex('ColB') == 1
    assert td.nrows() == 3
    assert td.ncols() == 2
    assert td.search ('ein') == [(0,1)]


