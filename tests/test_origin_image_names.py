"""Origin executable name detection from the COM registration
(origin2018-improvements).

Origin 2018 registers ``LocalServer32`` with an 8.3 short path, so tasklist
shows ``ORIGIN~1.EXE``; without reading the registry, PID capture silently
fails. A fake ``winreg`` module (installed in ``sys.modules``) stands in for
the Windows registry, so this runs on Linux too.
"""
import subprocess
import sys
import types

import pytest

from origin_pro_mcp import daemon

SHORT_83 = r"C:\PROGRA~1\ORIGIN~1\ORIGIN~1\ORIGIN~1.EXE"
QUOTED = '"C:\\Program Files\\X\\Origin64.exe" -Embedding'


def _fake_winreg(progid_clsid: dict, local_server: dict, *, view_only=None):
    """Build a fake winreg.

    progid_clsid: {"Origin.Application": "{CLSID}"}
    local_server: {"{CLSID}": "<LocalServer32 default value>"}
    view_only: if set, LocalServer32 is only visible under this access-view
        flag (models a key that exists only in the 32- or 64-bit view).
    """
    mod = types.ModuleType("winreg")
    mod.HKEY_CLASSES_ROOT = object()
    mod.KEY_READ = 0x20019
    mod.KEY_WOW64_32KEY = 0x0200
    mod.KEY_WOW64_64KEY = 0x0100
    mod.opened = []

    class _Key:
        def __init__(self, value):
            self.value = value

        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

    def OpenKey(root, sub, reserved=0, access=0):
        assert root is mod.HKEY_CLASSES_ROOT
        mod.opened.append((sub, access))
        if sub.endswith("\\CLSID") and not sub.startswith("CLSID\\"):
            progid = sub[: -len("\\CLSID")]
            if progid in progid_clsid:
                return _Key(progid_clsid[progid])
            raise FileNotFoundError(sub)
        if sub.startswith("CLSID\\") and sub.endswith("\\LocalServer32"):
            clsid = sub[len("CLSID\\"): -len("\\LocalServer32")]
            if clsid not in local_server:
                raise FileNotFoundError(sub)
            if view_only is not None and not (access & view_only):
                raise FileNotFoundError(sub)
            return _Key(local_server[clsid])
        raise FileNotFoundError(sub)

    def QueryValue(key, sub):
        assert sub is None
        return key.value

    mod.OpenKey = OpenKey
    mod.QueryValue = QueryValue
    return mod


@pytest.fixture
def no_override(monkeypatch):
    monkeypatch.delenv("ORIGIN_PRO_MCP_ORIGIN_IMAGE", raising=False)


# --- _registered_origin_images --------------------------------------------- #


def test_no_winreg_returns_empty_tuple(monkeypatch):
    # sys.modules[name] = None makes `import winreg` raise ImportError, which
    # is exactly what happens on Linux.
    monkeypatch.setitem(sys.modules, "winreg", None)
    assert daemon._registered_origin_images() == ()


def test_unquoted_83_short_path(monkeypatch):
    reg = _fake_winreg({"Origin.Application": "{A}"}, {"{A}": SHORT_83})
    monkeypatch.setitem(sys.modules, "winreg", reg)
    assert daemon._registered_origin_images() == ("ORIGIN~1.EXE",)


@pytest.mark.parametrize("suffix", [" -Embedding", " /Automation", "  "])
def test_unquoted_path_with_switches(monkeypatch, suffix):
    reg = _fake_winreg({"Origin.Application": "{A}"}, {"{A}": SHORT_83 + suffix})
    monkeypatch.setitem(sys.modules, "winreg", reg)
    assert daemon._registered_origin_images() == ("ORIGIN~1.EXE",)


def test_quoted_long_path_with_embedding(monkeypatch):
    reg = _fake_winreg({"Origin.Application": "{A}"}, {"{A}": QUOTED})
    monkeypatch.setitem(sys.modules, "winreg", reg)
    assert daemon._registered_origin_images() == ("Origin64.exe",)


def test_both_progids_are_read_and_deduplicated_case_insensitively(monkeypatch):
    reg = _fake_winreg(
        {"Origin.Application": "{A}", "Origin.ApplicationSI": "{B}"},
        {"{A}": SHORT_83, "{B}": r"C:\PROGRA~1\ORIGIN~1\ORIGIN~1\origin~1.exe -Embedding"},
    )
    monkeypatch.setitem(sys.modules, "winreg", reg)
    assert daemon._registered_origin_images() == ("ORIGIN~1.EXE",)


def test_two_different_servers_both_reported(monkeypatch):
    reg = _fake_winreg(
        {"Origin.Application": "{A}", "Origin.ApplicationSI": "{B}"},
        {"{A}": SHORT_83, "{B}": QUOTED},
    )
    monkeypatch.setitem(sys.modules, "winreg", reg)
    assert daemon._registered_origin_images() == ("ORIGIN~1.EXE", "Origin64.exe")


def test_missing_progid_or_server_key_is_skipped(monkeypatch):
    reg = _fake_winreg(
        {"Origin.Application": "{A}", "Origin.ApplicationSI": "{MISSING}"},
        {"{A}": QUOTED},
    )
    monkeypatch.setitem(sys.modules, "winreg", reg)
    assert daemon._registered_origin_images() == ("Origin64.exe",)


def test_nothing_registered_returns_empty(monkeypatch):
    monkeypatch.setitem(sys.modules, "winreg", _fake_winreg({}, {}))
    assert daemon._registered_origin_images() == ()


def test_key_only_in_64bit_view_is_found(monkeypatch):
    reg = _fake_winreg({"Origin.Application": "{A}"}, {"{A}": SHORT_83},
                       view_only=0x0100)  # KEY_WOW64_64KEY
    monkeypatch.setitem(sys.modules, "winreg", reg)
    assert daemon._registered_origin_images() == ("ORIGIN~1.EXE",)
    flags = [acc for sub, acc in reg.opened if sub.endswith("LocalServer32")]
    assert flags[-1] & 0x0100


def test_empty_server_value_is_ignored(monkeypatch):
    reg = _fake_winreg({"Origin.Application": "{A}"}, {"{A}": ""})
    monkeypatch.setitem(sys.modules, "winreg", reg)
    assert daemon._registered_origin_images() == ()


# --- _origin_image_names ---------------------------------------------------- #


def test_defaults_without_registry(monkeypatch, no_override):
    monkeypatch.setattr(daemon, "_registered_origin_images", lambda: ())
    assert daemon._origin_image_names() == daemon._DEFAULT_ORIGIN_IMAGES


def test_registry_name_appended_to_defaults(monkeypatch, no_override):
    monkeypatch.setattr(daemon, "_registered_origin_images", lambda: ("ORIGIN~1.EXE",))
    names = daemon._origin_image_names()
    assert names[: len(daemon._DEFAULT_ORIGIN_IMAGES)] == daemon._DEFAULT_ORIGIN_IMAGES
    assert names[-1] == "ORIGIN~1.EXE"


def test_registry_name_already_in_defaults_not_duplicated(monkeypatch, no_override):
    monkeypatch.setattr(daemon, "_registered_origin_images", lambda: ("ORIGIN64.EXE",))
    names = daemon._origin_image_names()
    assert names == daemon._DEFAULT_ORIGIN_IMAGES
    assert [n.lower() for n in names].count("origin64.exe") == 1


def test_override_replaces_defaults_but_registry_still_appended(monkeypatch):
    monkeypatch.setenv("ORIGIN_PRO_MCP_ORIGIN_IMAGE", " Custom.exe , Other.exe ,")
    monkeypatch.setattr(daemon, "_registered_origin_images", lambda: ("ORIGIN~1.EXE",))
    assert daemon._origin_image_names() == ("Custom.exe", "Other.exe", "ORIGIN~1.EXE")


def test_override_and_registry_dedup_is_case_insensitive(monkeypatch):
    monkeypatch.setenv("ORIGIN_PRO_MCP_ORIGIN_IMAGE", "origin~1.exe")
    monkeypatch.setattr(daemon, "_registered_origin_images",
                        lambda: ("ORIGIN~1.EXE", "Origin64.exe"))
    assert daemon._origin_image_names() == ("origin~1.exe", "Origin64.exe")


def test_end_to_end_fake_registry_feeds_tasklist_filter(monkeypatch, no_override):
    # tasklist lists every process; the registry name must be in the Python
    # filter so ORIGIN~1.EXE PIDs are captured (and unrelated ones are not).
    reg = _fake_winreg({"Origin.Application": "{A}"}, {"{A}": SHORT_83})
    monkeypatch.setitem(sys.modules, "winreg", reg)
    csv = (
        '"ORIGIN~1.EXE","31337","Console","1","250,000 K"\r\n'
        '"notepad.exe","4444","Console","1","10,000 K"\r\n'
    )
    monkeypatch.setattr(
        subprocess, "run",
        lambda cmd, *a, **k: subprocess.CompletedProcess(cmd, 0, stdout=csv, stderr=""),
    )
    assert daemon._origin_process_pids() == {31337}


def test_without_registry_short_name_is_missed(monkeypatch, no_override):
    # Documents the original Origin 2018 bug the registry lookup fixes.
    monkeypatch.setattr(daemon, "_registered_origin_images", lambda: ())
    csv = '"ORIGIN~1.EXE","31337","Console","1","250,000 K"\r\n'
    monkeypatch.setattr(
        subprocess, "run",
        lambda cmd, *a, **k: subprocess.CompletedProcess(cmd, 0, stdout=csv, stderr=""),
    )
    assert daemon._origin_process_pids() == set()
