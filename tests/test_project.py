from origin_connection import get_origin
import os

def test_new_project():
    o = get_origin()
    o.NewProject()

def test_save_and_load():
    o = get_origin()
    o.NewProject()
    o.CreatePage(2, "SaveTest", "origin")
    path = r"C:\Users\swym4\test_project.opju"
    o.Save(path)
    assert os.path.exists(path)
    # Close the project first so Origin releases the file lock, then delete
    o.NewProject()
    os.remove(path)
