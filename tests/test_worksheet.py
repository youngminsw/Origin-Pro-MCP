from origin_connection import get_origin

def setup_function():
    o = get_origin()
    o.Execute("doc -s; doc -n;")

def test_create_worksheet():
    o = get_origin()
    name = o.CreatePage(2, "TestWks", "origin")
    assert name == "TestWks"

def test_put_and_get_data():
    o = get_origin()
    o.CreatePage(2, "DataTest", "origin")
    o.PutWorksheet("[DataTest]Sheet1", [1.0, 2.0, 3.0], 0, 0)
    o.PutWorksheet("[DataTest]Sheet1", [4.0, 5.0, 6.0], 0, 1)
    data = o.GetWorksheet("[DataTest]Sheet1")
    assert data == ((1.0, 4.0), (2.0, 5.0), (3.0, 6.0))

def test_find_worksheet():
    o = get_origin()
    o.CreatePage(2, "ColTest", "origin")
    wks = o.FindWorksheet("[ColTest]Sheet1")
    assert wks is not None
