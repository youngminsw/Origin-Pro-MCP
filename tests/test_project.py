import os

import pytest

from origin_pro_mcp.origin_connection import get_origin

pytestmark = pytest.mark.requires_origin

def test_new_project():
    o = get_origin()
    o.NewProject()

def test_save_and_load():
    o = get_origin()
    o.NewProject()
    o.CreatePage(2, "SaveTest", "origin")
    path = os.path.join(os.path.expanduser("~"), "test_project.opju")
    o.Save(path)
    assert os.path.exists(path)
    # Close the project first so Origin releases the file lock, then delete
    o.NewProject()
    os.remove(path)
