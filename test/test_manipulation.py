from TableData import TableData

def test_delRow():     
    td=TableData.load_table ('test/data.xls')
    nrows=td.nrows()
    td.delRow(1)
    assert td.nrows() == nrows-1

def test_delRow():     
    td=TableData.load_table ('test/data.xls')
    ncols=td.ncols()
    td.delCol('ColA')
    assert td.ncols() == ncols-1

def test_addRow():
    td=TableData.load_table ('test/data.xls')
    ncols_old=td.ncols()
    td.addCol ('Maurice')
    assert td.ncols() == ncols_old+1
    assert td.cell(td.ncols()-1,0) == 'Maurice'
    #td.addCol ('Maurice')

def test_simple_things():
    td=TableData.load_table ('test/data.xls')
    assert td.colExists ('gibtsNicht') == False
    assert td.colExists ('ColA') == True
    
def test_addCol():
    td=TableData.load_table ('test/data.xls')
    cid=td.addCol('bla')
    assert cid == td.cindex('bla')