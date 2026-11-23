"""Write sandbox: ORIGIN_PRO_MCP_ALLOWED_ROOTS restricts where the server may
WRITE (exports, saves, templates, in-place autosave). Reading is never
restricted, and an unset variable keeps the old unrestricted behavior."""
from __future__ import annotations

import os

import pytest

from origin_pro_mcp import autosave, labtalk_safe
from origin_pro_mcp.labtalk_safe import (
    allowed_write_roots,
    is_writable_path,
    require_writable_path,
    writable_path,
)

ENV = "ORIGIN_PRO_MCP_ALLOWED_ROOTS"


@pytest.fixture
def roots(monkeypatch, tmp_path):
    allowed = tmp_path / "allowed"
    allowed.mkdir()
    outside = tmp_path / "outside"
    outside.mkdir()
    monkeypatch.setenv(ENV, str(allowed))
    return allowed, outside


@pytest.fixture(autouse=True)
def _clear(monkeypatch):
    monkeypatch.delenv(ENV, raising=False)


# --------------------------------------------------------------------------- #
# Parsing and path checks                                                     #
# --------------------------------------------------------------------------- #


def test_unset_means_unrestricted(tmp_path):
    assert allowed_write_roots() == []
    assert is_writable_path(str(tmp_path / "anything.csv"))
    assert require_writable_path("/any/where.png", "file_path") == "/any/where.png"


def test_roots_split_on_semicolons_and_newlines_and_strip_quotes(monkeypatch):
    monkeypatch.setenv(ENV, ' "C:\\A" ; C:\\B\n\nD:\\C ;; ')
    assert allowed_write_roots() == ["C:\\A", "C:\\B", "D:\\C"]


def test_colon_is_not_a_separator(monkeypatch):
    monkeypatch.setenv(ENV, "C:\\Data\\out")
    assert allowed_write_roots() == ["C:\\Data\\out"]


def test_inside_root_and_root_itself_allowed(roots):
    allowed, _ = roots
    assert is_writable_path(str(allowed))
    assert is_writable_path(str(allowed / "fig.png"))
    assert is_writable_path(str(allowed / "sub" / "deeper" / "x.csv"))


def test_outside_root_rejected(roots):
    _, outside = roots
    assert not is_writable_path(str(outside / "fig.png"))


def test_sibling_with_common_prefix_rejected(roots, tmp_path):
    # "allowed2" starts with "allowed" but is NOT inside it.
    assert not is_writable_path(str(tmp_path / "allowed2" / "x.png"))


def test_dotdot_escape_rejected(roots):
    allowed, _ = roots
    assert not is_writable_path(str(allowed / ".." / "outside" / "x.png"))


def test_multiple_roots(monkeypatch, tmp_path):
    a, b = tmp_path / "a", tmp_path / "b"
    monkeypatch.setenv(ENV, f"{a};{b}")
    assert is_writable_path(str(a / "1.png")) and is_writable_path(str(b / "2.png"))
    assert not is_writable_path(str(tmp_path / "c" / "3.png"))


@pytest.mark.skipif(os.name != "nt", reason="case-insensitive paths are Windows-only")
def test_case_insensitive_on_windows(roots):
    allowed, _ = roots
    assert is_writable_path(str(allowed).upper() + "\\x.png")


def test_different_drive_is_outside(monkeypatch):
    if os.name != "nt":
        pytest.skip("drive letters are Windows-only")
    monkeypatch.setenv(ENV, "C:\\allowed")
    assert not is_writable_path("D:\\allowed\\x.png")


def test_error_message_names_the_setting_and_says_reading_is_fine(roots):
    _, outside = roots
    with pytest.raises(ValueError) as exc:
        require_writable_path(str(outside / "x.png"), "file_path")
    msg = str(exc.value)
    assert ENV in msg and "outside the folders" in msg and "Reading" in msg


def test_writable_path_normalizes_then_checks(roots, monkeypatch):
    allowed, outside = roots
    monkeypatch.setattr(labtalk_safe, "windows_path", lambda v, f: v.strip('"'))
    assert writable_path(f'"{allowed / "x.png"}"', "file_path") == str(allowed / "x.png")
    with pytest.raises(ValueError, match="outside the folders"):
        writable_path(str(outside / "x.png"), "file_path")


# --------------------------------------------------------------------------- #
# Tools refuse BEFORE writing                                                 #
# --------------------------------------------------------------------------- #


def test_export_worksheet_outside_root_writes_nothing(fake_origin, roots):
    from origin_pro_mcp.tools.worksheet import export_worksheet

    _, outside = roots
    target = outside / "out.csv"
    with pytest.raises(ValueError, match="outside the folders"):
        export_worksheet("Book1", "Sheet1", str(target))
    assert not target.exists()


def test_export_worksheet_inside_root_works(fake_origin, roots):
    from origin_pro_mcp.tools.worksheet import export_worksheet

    allowed, _ = roots
    target = allowed / "out.csv"
    export_worksheet("Book1", "Sheet1", str(target))
    assert target.exists()


def test_export_graph_outside_root_refused(fake_origin, roots):
    from origin_pro_mcp.tools.graph import export_graph

    _, outside = roots
    with pytest.raises(ValueError, match="outside the folders"):
        export_graph("Graph1", str(outside / "fig.png"))
    assert not (outside / "fig.png").exists()


def test_export_all_graphs_outside_root_creates_no_folder(fake_origin, roots):
    from origin_pro_mcp.tools.project import export_all_graphs

    _, outside = roots
    new_dir = outside / "figs"
    with pytest.raises(ValueError, match="outside the folders"):
        export_all_graphs(str(new_dir))
    assert not new_dir.exists()


def test_save_graph_template_outside_root_refused(fake_origin, roots):
    from origin_pro_mcp.tools.project import save_graph_template

    _, outside = roots
    with pytest.raises(ValueError, match="outside the folders"):
        save_graph_template("Graph1", str(outside / "t.otpu"))


def test_save_project_explicit_path_outside_root_refused(fake_origin, roots):
    from origin_pro_mcp.tools.project import save_project

    _, outside = roots
    with pytest.raises(ValueError, match="outside the folders"):
        save_project(str(outside / "p.opju"))
    assert fake_origin.saved_paths == []


def test_save_project_in_place_outside_root_refused(fake_origin, roots):
    from origin_pro_mcp import origin_connection
    from origin_pro_mcp.tools.project import save_project

    _, outside = roots
    original = outside / "original.opju"
    original.write_bytes(b"REAL" * 1000)
    origin_connection.remember_project_path(str(original))
    with pytest.raises(ValueError, match="outside the folders"):
        save_project()
    assert fake_origin.saved_paths == []


def test_save_project_in_place_unknown_location_refused_when_restricted(fake_origin, roots, monkeypatch):
    from origin_pro_mcp.tools import project

    monkeypatch.setattr(project, "_current_project_location", lambda o: None)
    with pytest.raises(ValueError, match="location is unknown"):
        project.save_project()


def test_save_project_in_place_unrestricted_still_calls_save(fake_origin, monkeypatch):
    from origin_pro_mcp.tools import project

    monkeypatch.setattr(project, "_current_project_location", lambda o: None)
    # No roots configured -> the location check is skipped entirely.
    project._require_current_location_writable(fake_origin, "save in place")


# --------------------------------------------------------------------------- #
# Autosave never overwrites a file outside the roots                          #
# --------------------------------------------------------------------------- #


class _AutoOrigin:
    def __init__(self):
        self.saved = []
        count = type("C", (), {"Count": 1})()
        self.WorksheetPages = count
        self.GraphPages = count
        self.MatrixPages = count

    def Save(self, path):
        self.saved.append(path)
        return True


def test_autosave_skips_original_outside_roots(roots):
    _, outside = roots
    original = outside / "original.opju"
    original.write_bytes(b"REAL" * 1000)
    o = _AutoOrigin()
    assert autosave.save_in_place(o, str(original)) is None
    assert o.saved == []


def test_autosave_inside_roots_still_saves(roots):
    allowed, _ = roots
    copy = allowed / "copy.opju"
    copy.write_bytes(b"REAL" * 1000)
    o = _AutoOrigin()
    assert autosave.save_in_place(o, str(copy)) is True
    assert o.saved == [str(copy)]


# --------------------------------------------------------------------------- #
# load_project writes its pre-load .bak only inside the roots                 #
# --------------------------------------------------------------------------- #


def test_load_outside_roots_writes_no_backup(roots, fake_origin):
    from origin_pro_mcp.tools.project import load_project

    _, outside = roots
    original = outside / "original.opju"
    original.write_bytes(b"REAL" * 1000)
    load_project(str(original))                     # opening is never restricted
    assert not (outside / "original.opju.bak").exists()
    assert original.read_bytes() == b"REAL" * 1000


def test_load_inside_roots_still_backs_up(roots, fake_origin):
    from origin_pro_mcp.tools.project import load_project

    allowed, _ = roots
    copy = allowed / "copy.opju"
    copy.write_bytes(b"REAL" * 1000)
    load_project(str(copy))
    assert (allowed / "copy.opju.bak").read_bytes() == b"REAL" * 1000
