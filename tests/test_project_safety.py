"""N5 (data destruction): save_project must never let an EMPTY project overwrite
a real .opju, must keep a pre-overwrite .bak, and load_project must keep a
pre-load .bak (o.Load is read-only, but the backup is the recovery net)."""
import pytest


def _make_empty(fake):
    fake.books = []
    fake.graphs = []
    fake.matrices = []


def test_save_project_refuses_empty_over_real_file(fake_origin, tmp_path):
    from origin_pro_mcp.tools.project import save_project
    _make_empty(fake_origin)                     # 0 windows (e.g. after flaky load)
    real = tmp_path / "study.opju"
    real.write_bytes(b"REALPROJECT" * 500)       # ~5.5 KB "real" project
    before = real.read_bytes()
    with pytest.raises(ValueError, match="Refusing to save"):
        save_project(str(real))
    assert real.read_bytes() == before           # untouched
    assert fake_origin.saved_paths == []         # Origin.Save never called


def test_save_project_no_path_refuses_when_empty(fake_origin):
    from origin_pro_mcp.tools.project import save_project
    _make_empty(fake_origin)
    with pytest.raises(ValueError, match="Refusing to save"):
        save_project()                            # save-to-current with empty project
    assert fake_origin.saved_paths == []


def test_save_project_backs_up_before_overwrite(fake_origin, tmp_path):
    from origin_pro_mcp.tools.project import save_project
    real = tmp_path / "study.opju"
    real.write_bytes(b"OLD" * 1000)               # existing project to be overwritten
    out = save_project(str(real))                 # non-empty fake -> saves + backs up
    assert (tmp_path / "study.opju.bak").exists()
    assert (tmp_path / "study.opju.bak").read_bytes() == b"OLD" * 1000
    assert "backup" in out
    assert str(real) in fake_origin.saved_paths


def test_save_project_no_backup_for_new_file(fake_origin, tmp_path):
    from origin_pro_mcp.tools.project import save_project
    target = tmp_path / "brand_new.opju"          # does not exist yet
    save_project(str(target))
    assert not (tmp_path / "brand_new.opju.bak").exists()


def test_load_project_backs_up_target(fake_origin, tmp_path):
    from origin_pro_mcp.tools.project import load_project
    real = tmp_path / "study.opju"
    real.write_bytes(b"DATA" * 1000)
    load_project(str(real))                       # non-empty fake -> load succeeds
    assert (tmp_path / "study.opju.bak").exists()
    assert (tmp_path / "study.opju.bak").read_bytes() == b"DATA" * 1000
