"""Guard tests for data IO / export tools (no Origin needed)."""
import json

import pytest


def test_export_worksheet_bad_delimiter(fake_origin):
    from origin_pro_mcp.tools.worksheet import export_worksheet

    with pytest.raises(ValueError, match="delimiter must be one of"):
        export_worksheet("Book1", "Sheet1", "/tmp/x.csv", delimiter="::")


def test_export_worksheet_unknown_sheet(fake_origin):
    from origin_pro_mcp.tools.worksheet import export_worksheet

    with pytest.raises(ValueError, match="not found"):
        export_worksheet("Ghost", "Sheet1", "/tmp/x.csv")


def test_export_worksheet_writes_csv(fake_origin, tmp_path):
    from origin_pro_mcp.tools.worksheet import export_worksheet

    # FakeOrigin.worksheet_data is ((1,4),(2,5)); ncols defaults to 0 so
    # the header is empty, but rows must still be written.
    out = tmp_path / "out.csv"
    msg = export_worksheet("Book1", "Sheet1", str(out))
    assert out.exists()
    assert "rows" in msg
    assert out.read_text().strip().splitlines()[-1] == "2.0,5.0"


def test_import_excel_missing_file(fake_origin, tmp_path):
    from origin_pro_mcp.tools.worksheet import import_data

    # format="auto" routes .xlsx through the Excel impl.
    with pytest.raises(ValueError, match="File not found"):
        import_data(str(tmp_path / "nope.xlsx"))


def test_import_data_csv_returns_json_name(fake_origin, tmp_path):
    from origin_pro_mcp.tools.worksheet import import_data

    f = tmp_path / "data.csv"
    f.write_text("1,2\n3,4\n")
    fake_origin.LTStr = lambda name: "Book2" if name == "page.name$" else ""
    out = json.loads(import_data(str(f), book_name="Book2"))
    assert out["name"] == "Book2"
    assert out["requested_name"] == "Book2"
    assert out["renamed"] is False
    assert out["file"] == str(f)


def test_import_data_csv_no_book_name_requested_is_null(fake_origin, tmp_path):
    from origin_pro_mcp.tools.worksheet import import_data

    f = tmp_path / "data.csv"
    f.write_text("1,2\n")
    fake_origin.LTStr = lambda name: "Book1" if name == "page.name$" else ""
    out = json.loads(import_data(str(f)))
    assert out["name"] == "Book1"
    assert out["requested_name"] is None
    assert out["renamed"] is False


def test_import_data_excel_returns_json_name(fake_origin, tmp_path):
    from origin_pro_mcp.tools.worksheet import import_data

    f = tmp_path / "data.xlsx"
    f.write_bytes(b"PK\x03\x04stub")
    fake_origin.LTStr = lambda name: "Book1" if name == "page.name$" else ""
    out = json.loads(import_data(str(f)))
    assert out["name"] == "Book1"
    assert out["requested_name"] is None
    assert out["renamed"] is False
    assert out["file"] == str(f)


def test_export_graph_sized_unknown_graph(fake_origin, tmp_path):
    from origin_pro_mcp.tools.graph import export_graph

    with pytest.raises(ValueError, match="not found"):
        export_graph("Ghost", str(tmp_path / "g.png"), sized=True)

def test_export_graph_reports_uncreated_file_with_wsl_hint(fake_origin, tmp_path):
    from origin_pro_mcp.tools.graph import export_graph

    # FakeOrigin's Execute "succeeds" but no file is written: the guard must
    # turn that into a loud failure (not a stale/silent success), and since the
    # test path is POSIX it must also hint that Origin (Windows) can't reach it.
    with pytest.raises(ValueError, match="was not created.*WSL/Linux path"):
        export_graph("Graph1", str(tmp_path / "fig.png"))


def test_save_graph_template_unknown_graph(fake_origin, tmp_path):
    from origin_pro_mcp.tools.project import save_graph_template

    with pytest.raises(ValueError, match="not found"):
        save_graph_template("Ghost", str(tmp_path / "t.otpu"))


def test_save_graph_template_forces_extension(fake_origin, tmp_path):
    from origin_pro_mcp.tools.project import save_graph_template

    # FakeOrigin doesn't write the file, so it raises after the save attempt;
    # assert the executed command used the .otpu extension and save -t.
    with pytest.raises(ValueError, match="was not created"):
        save_graph_template("Graph1", str(tmp_path / "mytmpl.png"))
    assert any(
        s.startswith("save -t Graph1") and "mytmpl.otpu" in s
        for s in fake_origin.executed
    )
