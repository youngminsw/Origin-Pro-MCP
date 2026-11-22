"""get_worksheet_data paging, column labels and missing-value handling
(origin2018-improvements). COM-free via the shared FakeOrigin."""
import json

import pytest
from fakes import FakeBook, FakeColumn, FakeSheet

from origin_pro_mcp.tools.worksheet import (
    GET_WORKSHEET_DEFAULT_MAX_ROWS,
    get_worksheet_data,
)


def _rows(n, ncols=2):
    # Row i -> (i, 1000+i, ...): easy to tell which rows came back.
    return tuple(tuple(float(c * 1000 + i) for c in range(ncols)) for i in range(n))


def _read(**kw):
    return json.loads(get_worksheet_data("Book1", "Sheet1", **kw))


def test_default_page_is_200_rows(fake_origin):
    assert GET_WORKSHEET_DEFAULT_MAX_ROWS == 200
    fake_origin.worksheet_data = _rows(450)
    out = _read()
    assert out["total_rows"] == 450
    assert out["total_cols"] == 2
    assert out["start_row"] == 0
    assert out["returned_rows"] == 200
    assert len(out["columns"]) == 2
    assert out["columns"][0] == [float(i) for i in range(200)]
    assert out["columns"][1][0] == 1000.0
    assert out["truncated"] is True
    assert out["next_start_row"] == 200


def test_start_row_and_max_rows_select_a_slice(fake_origin):
    fake_origin.worksheet_data = _rows(450)
    out = _read(start_row=200, max_rows=50)
    assert out["columns"][0] == [float(i) for i in range(200, 250)]
    assert out["returned_rows"] == 50
    assert out["truncated"] is True
    assert out["next_start_row"] == 250


def test_last_page_is_not_truncated(fake_origin):
    fake_origin.worksheet_data = _rows(450)
    out = _read(start_row=400)
    assert out["columns"][0] == [float(i) for i in range(400, 450)]
    assert out["returned_rows"] == 50
    assert out["truncated"] is False
    assert out["next_start_row"] is None


def test_paging_walks_whole_sheet_without_gaps(fake_origin):
    fake_origin.worksheet_data = _rows(450)
    seen, start = [], 0
    while start is not None:
        out = _read(start_row=start, max_rows=200)
        seen.extend(out["columns"][0])
        start = out["next_start_row"]
    assert seen == [float(i) for i in range(450)]


def test_max_rows_zero_returns_everything(fake_origin):
    fake_origin.worksheet_data = _rows(450)
    out = _read(max_rows=0)
    assert out["returned_rows"] == 450
    assert len(out["columns"][0]) == 450
    assert out["truncated"] is False
    assert out["next_start_row"] is None


def test_max_rows_zero_with_start_row_returns_the_rest(fake_origin):
    fake_origin.worksheet_data = _rows(10)
    out = _read(start_row=7, max_rows=0)
    assert out["columns"][0] == [7.0, 8.0, 9.0]
    assert out["truncated"] is False


def test_small_sheet_fits_in_one_page(fake_origin):
    out = _read()  # default fake: 2 rows x 2 cols
    assert out["columns"] == [[1.0, 2.0], [4.0, 5.0]]
    assert out["total_rows"] == 2
    assert out["truncated"] is False
    assert out["next_start_row"] is None


def test_start_row_equal_to_total_is_an_empty_page(fake_origin):
    fake_origin.worksheet_data = _rows(5)
    out = _read(start_row=5)
    assert out["returned_rows"] == 0
    assert out["columns"] == [[], []]
    assert out["truncated"] is False


def test_start_row_past_end_raises(fake_origin):
    fake_origin.worksheet_data = _rows(5)
    with pytest.raises(ValueError, match="past the end") as exc:
        _read(start_row=6)
    assert "5 rows" in str(exc.value)


@pytest.mark.parametrize("bad", [-1, 1.5, "3", True])
def test_invalid_start_row_raises(fake_origin, bad):
    with pytest.raises(ValueError, match="start_row"):
        _read(start_row=bad)


@pytest.mark.parametrize("bad", [-1, 2.0, "10", False])
def test_invalid_max_rows_raises(fake_origin, bad):
    with pytest.raises(ValueError, match="max_rows"):
        _read(max_rows=bad)


def test_paging_args_validated_before_touching_origin(fake_origin):
    # A bad argument must not even issue a GetWorksheet round trip.
    calls = []
    fake_origin.GetWorksheet = lambda target: calls.append(target)
    with pytest.raises(ValueError):
        _read(start_row=-5)
    assert calls == []


def test_empty_sheet(fake_origin):
    fake_origin.worksheet_data = ()
    out = _read()
    assert out["columns"] == []
    assert out["total_rows"] == 0
    assert out["total_cols"] == 0
    assert out["long_names"] == []
    assert out["units"] == []
    assert out["truncated"] is False
    assert out["next_start_row"] is None


def test_empty_sheet_start_row_past_end_raises(fake_origin):
    fake_origin.worksheet_data = ()
    with pytest.raises(ValueError, match="past the end"):
        _read(start_row=1)


# --- missing values -------------------------------------------------------- #


def test_missing_value_sentinel_becomes_null(fake_origin):
    fake_origin.worksheet_data = (
        (1.0, 1.23e308),
        (-1.7e308, 5.0),
        (1e100, 9.9e99),   # threshold is >= 1e100 in magnitude
        ("txt", 0.0),      # non-numeric cells pass through untouched
    )
    out = _read()
    assert out["columns"] == [[1.0, None, None, "txt"], [None, 5.0, 9.9e99, 0.0]]


def test_missing_value_mapping_applies_inside_a_page(fake_origin):
    data = list(_rows(300))
    data[250] = (1.23e308, 7.0)
    fake_origin.worksheet_data = tuple(data)
    out = _read(start_row=250, max_rows=2)
    assert out["columns"][0] == [None, 251.0]
    assert out["columns"][1] == [7.0, 1251.0]


# --- long names / units ---------------------------------------------------- #


class _UnitColumn(FakeColumn):
    def __init__(self, name, long_name="", units=""):
        super().__init__(name, long_name=long_name)
        self.Units = units


def test_long_names_and_units_are_returned(fake_origin):
    sheet = FakeSheet("Sheet1", columns=[
        _UnitColumn("A", long_name="Time", units="s"),
        _UnitColumn("B", long_name="Voltage", units="V"),
    ])
    fake_origin.books = [FakeBook("Book1", sheets=(sheet,))]
    out = _read()
    assert out["long_names"] == ["Time", "Voltage"]
    assert out["units"] == ["s", "V"]


def test_missing_units_attribute_yields_empty_string(fake_origin):
    # FakeColumn has LongName but no Units -> getattr raises AttributeError.
    sheet = FakeSheet("Sheet1", columns=[FakeColumn("A", long_name="Time"),
                                         FakeColumn("B")])
    fake_origin.books = [FakeBook("Book1", sheets=(sheet,))]
    out = _read()
    assert out["long_names"] == ["Time", ""]
    assert out["units"] == ["", ""]


class _ExplodingColumn:
    @property
    def LongName(self):
        raise OSError("COM error reading LongName")

    @property
    def Units(self):
        raise OSError("COM error reading Units")


class _ExplodingColumns:
    Count = 2

    def Item(self, i):
        if i == 0:
            return _ExplodingColumn()
        raise OSError("COM error on Item")


class _ExplodingSheet:
    Name = "Sheet1"

    @property
    def Columns(self):
        return _ExplodingColumns()


def test_label_com_errors_do_not_fail_the_read(fake_origin):
    fake_origin.FindWorksheet = lambda target: _ExplodingSheet()
    fake_origin.GetWorksheet = lambda target: ((1.0, 2.0),)
    out = _read()
    assert out["columns"] == [[1.0], [2.0]]
    assert out["long_names"] == ["", ""]
    assert out["units"] == ["", ""]


def test_find_worksheet_raising_does_not_fail_the_read(fake_origin):
    def boom(target):
        raise OSError("RPC server unavailable")

    fake_origin.FindWorksheet = boom
    fake_origin.GetWorksheet = lambda target: ((1.0, 2.0),)
    out = _read()
    assert out["columns"] == [[1.0], [2.0]]
    assert out["long_names"] == ["", ""]
    assert out["units"] == ["", ""]


def test_non_string_label_values_become_empty(fake_origin):
    sheet = FakeSheet("Sheet1", columns=[_UnitColumn("A", long_name=None, units=3),
                                         _UnitColumn("B", long_name="ok", units="m")])
    fake_origin.books = [FakeBook("Book1", sheets=(sheet,))]
    out = _read()
    assert out["long_names"] == ["", "ok"]
    assert out["units"] == ["", "m"]


def test_fewer_com_columns_than_data_columns_pads_labels(fake_origin):
    sheet = FakeSheet("Sheet1", columns=[_UnitColumn("A", long_name="T", units="s")])
    fake_origin.books = [FakeBook("Book1", sheets=(sheet,))]
    out = _read()  # data still has 2 columns
    assert out["long_names"] == ["T", ""]
    assert out["units"] == ["s", ""]
