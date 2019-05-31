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