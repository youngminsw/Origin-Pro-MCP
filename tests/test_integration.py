"""End-to-end test: create data, plot, style, fit, export."""
import win32com.client
from origin_connection import get_origin, execute_labtalk, get_lt_var
from PIL import ImageGrab
import os

def test_full_workflow():
    o = win32com.client.gencache.EnsureDispatch("Origin.ApplicationSI")
    o.NewProject()

    # 1. Create workbook with data
    o.CreatePage(2, "Experiment", "origin")
    x = [float(i) for i in range(1, 11)]
    y = [2.1*i + 0.5 + (i%3)*0.1 for i in range(1, 11)]
    o.PutWorksheet("[Experiment]Sheet1", x, 0, 0)
    o.PutWorksheet("[Experiment]Sheet1", y, 0, 1)

    # 2. Create graph
    o.CreatePage(3, "Figure1", "origin")
    o.Execute("plotxy iy:=[Experiment]Sheet1!(1,2) plot:=202 ogl:=[Figure1]Layer1;")

    # 3. Style for publication
    o.Execute('win -a Figure1;')
    o.Execute('xb.text$ = "Time (s)";')
    o.Execute('yl.text$ = "Signal (mV)";')
    o.Execute('xb.fsize = 12;')
    o.Execute('yl.fsize = 12;')

    # 4. Export via CopyPage + clipboard (expGraph requires interactive window focus)
    out = r"C:\Users\swym4\test_integration_fig.png"
    # CopyPage format=4 gives high-res BMP copy to clipboard
    o.CopyPage("Figure1", 4, 96, 24)
    img = ImageGrab.grabclipboard()
    assert img is not None, "CopyPage clipboard grab returned None"
    img.save(out)

    assert os.path.exists(out), "Export failed: file not created"
    size = os.path.getsize(out)
    assert size > 1000, f"Export file too small: {size} bytes"
    print(f"Integration test passed! Exported {size} bytes")
    os.remove(out)
