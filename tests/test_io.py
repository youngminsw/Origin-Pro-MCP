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


# --- item 31b: batch/folder import -------------------------------------------

def test_import_data_batch_directory(fake_origin, tmp_path):
    from origin_pro_mcp.tools.worksheet import import_data

    for i in range(3):
        (tmp_path / f"spectrum{i}.csv").write_text("1,2\n3,4\n")
    out = json.loads(import_data(str(tmp_path)))
    assert out["batch"] is True
    assert out["matched"] == 3
    assert out["imported"] == 3
    assert len(out["results"]) == 3
    assert all(r["ok"] for r in out["results"])


def test_import_data_batch_glob(fake_origin, tmp_path):
    from origin_pro_mcp.tools.worksheet import import_data

    (tmp_path / "a.csv").write_text("1,2\n")
    (tmp_path / "b.csv").write_text("1,2\n")
    (tmp_path / "skip.log").write_text("nope\n")  # not a data extension
    out = json.loads(import_data(str(tmp_path / "*.csv")))
    assert out["matched"] == 2
    assert {r["file"].split("/")[-1].split("\\")[-1] for r in out["results"]} == {"a.csv", "b.csv"}


def test_import_data_batch_no_matches_raises(fake_origin, tmp_path):
    from origin_pro_mcp.tools.worksheet import import_data

    with pytest.raises(ValueError, match="No data files"):
        import_data(str(tmp_path))


def test_import_data_batch_caps_at_20(fake_origin, tmp_path):
    from origin_pro_mcp.tools.worksheet import import_data

    for i in range(23):
        (tmp_path / f"f{i:02d}.csv").write_text("1,2\n")
    out = json.loads(import_data(str(tmp_path)))
    assert out["matched"] == 23
    assert out["imported"] == 20
    assert len(out["results"]) == 20
    assert "note" in out and "20" in out["note"]


def test_import_data_batch_reports_per_file_failure(fake_origin, tmp_path, monkeypatch):
    # A per-file import error is reported in that file's result, not fatal.
    import origin_pro_mcp.tools.worksheet as W

    (tmp_path / "good.csv").write_text("1,2\n")
    (tmp_path / "bad.csv").write_text("1,2\n")
    real = W._import_csv_to_worksheet_impl

    def flaky(path, book_name, delimiter, sparklines):
        if "bad" in path:
            raise ValueError("boom")
        return real(path, book_name, delimiter, sparklines)

    monkeypatch.setattr(W, "_import_csv_to_worksheet_impl", flaky)
    out = json.loads(W.import_data(str(tmp_path)))
    assert out["matched"] == 2
    assert out["imported"] == 1
    by_ok = {r["ok"]: r for r in out["results"]}
    assert "boom" in by_ok[False]["error"]


def test_book_name_from_stem_sanitizes():
    from origin_pro_mcp.tools.worksheet import _book_name_from_stem

    assert _book_name_from_stem("2024 run-1") == "B_2024_run_1"
    assert _book_name_from_stem("clean_name") == "clean_name"


def test_import_data_csv_activates_uniquified_book_not_existing(fake_origin, tmp_path):
    """Item 3: when book_name collides, CreatePage uniquifies it. The import
    must activate (and land in) the NEW uniquified book, never `win -a` the
    pre-existing book of the requested name."""
    from origin_pro_mcp.tools.worksheet import import_data

    f = tmp_path / "data.csv"
    f.write_text("1,2\n3,4\n")
    # Model CreatePage uniquifying "Data" -> "Data1" (the requested name is taken).
    fake_origin.CreatePage = lambda kind, name, tmpl: "Data1"
    fake_origin.LTStr = lambda name: "Data1" if name == "page.name$" else ""

    out = json.loads(import_data(str(f), book_name="Data"))
    assert out["name"] == "Data1"
    assert out["requested_name"] == "Data"
    assert out["renamed"] is True
    # Activated the actual new book, never the pre-existing "Data".
    assert any(s.startswith("win -a Data1") for s in fake_origin.executed)
    assert not any(s.strip() == "win -a Data;" for s in fake_origin.executed)


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


def test_import_data_csv_sparklines_default_deletes_new_graph_windows(fake_origin, tmp_path, monkeypatch):
    from origin_pro_mcp.tools.worksheet import import_data
    from fakes import FakeGraph

    f = tmp_path / "data.csv"
    f.write_text("1,2\n")
    fake_origin.LTStr = lambda name: "Book1" if name == "page.name$" else ""

    original_execute = fake_origin.Execute

    def fake_execute(script):
        if script.startswith("impasc"):
            fake_origin.graphs.append(FakeGraph("Spark1"))
            fake_origin.graphs.append(FakeGraph("Spark2"))
        return original_execute(script)

    monkeypatch.setattr(fake_origin, "Execute", fake_execute)

    out = json.loads(import_data(str(f)))
    # Item 16 (usability F12): the option "ran" but 2 sparkline windows still
    # leaked and had to be cleaned up, so suppression did NOT actually work —
    # the two fields must not contradict (suppressed cannot be True here).
    assert out["sparklines_suppressed"] is False
    assert out["sparklines_deleted"] == 2
    assert "win -cd Spark1;" in fake_origin.executed
    assert "win -cd Spark2;" in fake_origin.executed
    # Pre-existing windows must never be touched.
    assert not any(s == "win -cd Graph1;" for s in fake_origin.executed)


def test_import_data_csv_sparklines_suppressed_when_no_windows_leak(fake_origin, tmp_path):
    """Item 16: with the correct options.Sparklines:=0 key the import produces
    NO sparkline windows, so suppression is reported True and nothing is
    cleaned up — the non-contradictory happy path."""
    from origin_pro_mcp.tools.worksheet import import_data

    f = tmp_path / "data.csv"
    f.write_text("1,2\n")
    fake_origin.LTStr = lambda name: "Book1" if name == "page.name$" else ""

    out = json.loads(import_data(str(f)))
    assert out["sparklines_suppressed"] is True
    assert out["sparklines_deleted"] == 0
    # The correct key is issued.
    assert any("options.Sparklines:=0" in s for s in fake_origin.executed)


def test_import_data_csv_sparklines_true_skips_cleanup(fake_origin, tmp_path, monkeypatch):
    from origin_pro_mcp.tools.worksheet import import_data
    from fakes import FakeGraph

    f = tmp_path / "data.csv"
    f.write_text("1,2\n")
    fake_origin.LTStr = lambda name: "Book1" if name == "page.name$" else ""

    original_execute = fake_origin.Execute

    def fake_execute(script):
        if script.startswith("impasc"):
            fake_origin.graphs.append(FakeGraph("Spark1"))
        return original_execute(script)

    monkeypatch.setattr(fake_origin, "Execute", fake_execute)

    out = json.loads(import_data(str(f), sparklines=True))
    assert out["sparklines_suppressed"] is False
    assert out["sparklines_deleted"] == 0
    assert not any(s.startswith("win -cd") for s in fake_origin.executed)


def test_import_data_csv_sparkline_option_unsupported_falls_back(fake_origin, tmp_path, monkeypatch):
    from origin_pro_mcp.tools.worksheet import import_data

    f = tmp_path / "data.csv"
    f.write_text("1,2\n")
    fake_origin.LTStr = lambda name: "Book1" if name == "page.name$" else ""

    def fake_execute(script):
        fake_origin.executed.append(script)
        return "Sparklines" not in script

    monkeypatch.setattr(fake_origin, "Execute", fake_execute)

    out = json.loads(import_data(str(f)))
    assert out["sparklines_suppressed"] is False
    impasc_calls = [s for s in fake_origin.executed if s.startswith("impasc")]
    assert len(impasc_calls) == 2
    assert "Sparklines" in impasc_calls[0]
    assert "Sparklines" not in impasc_calls[1]


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
