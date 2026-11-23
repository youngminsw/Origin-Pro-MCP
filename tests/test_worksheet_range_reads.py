"""get_worksheet_data reads one page with GetWorksheet row ranges.

A page used to transfer the whole sheet over COM and slice it in Python. It
now asks Origin only for the page (plus one row to learn whether more
follow), and counts the total with single-row reads only when it has to.
The fake's range semantics follow what was live-checked on Origin 2018:
r1..r2 inclusive, r2 past the end clipped, r1 == rows -> empty, r1 > rows
-> None, and an existing EMPTY sheet answers with an HRESULT int.
"""
import json
import math

import pytest

from origin_pro_mcp.tools.worksheet import get_worksheet_data

HRESULT = -2147352568


def _rows(n):
    return tuple((float(i), float(1000 + i)) for i in range(n))


def _read(**kw):
    return json.loads(get_worksheet_data("Book1", "Sheet1", **kw))


def _expected(n, start, max_rows):
    data = _rows(n)
    end = n if max_rows == 0 else min(n, start + max_rows)
    page = data[start:end]
    return {
        "columns": [[r[c] for r in page] for c in range(2)] if n else [],
        "total_rows": n,
        "total_cols": 2 if n else 0,
        "returned_rows": len(page),
        "truncated": end < n,
        "next_start_row": end if end < n else None,
    }


CASES = [
    (n, start, max_rows)
    for n in (1, 2, 5, 199, 200, 201, 202, 450, 1000, 1024, 1025)
    for max_rows in (1, 7, 200)
    for start in sorted({0, n // 3, max(0, n - max_rows), n - 1, n})
]


@pytest.mark.parametrize("n,start,max_rows", CASES)
def test_page_matches_a_whole_sheet_slice(fake_origin, n, start, max_rows):
    fake_origin.worksheet_data = _rows(n)
    out = _read(start_row=start, max_rows=max_rows)
    for key, value in _expected(n, start, max_rows).items():
        assert out[key] == value, (key, n, start, max_rows)


def test_a_page_never_reads_the_whole_sheet(fake_origin):
    n = 100_000
    fake_origin.worksheet_data = _rows(n)
    out = _read(start_row=500, max_rows=200)
    assert out["total_rows"] == n and out["returned_rows"] == 200
    calls = fake_origin.get_worksheet_calls
    assert all(r1 is not None for _, r1, *_ in calls), "no full-sheet read"
    # One page read plus a logarithmic number of single-row probes.
    assert len(calls) <= 2 * math.ceil(math.log2(n)) + 3, len(calls)
    page_reads = [c for c in calls if c[3] - c[1] > 0]
    assert len(page_reads) == 1 and page_reads[0][1:4] == (500, 0, 700)


def test_last_page_needs_no_row_count(fake_origin):
    fake_origin.worksheet_data = _rows(450)
    out = _read(start_row=400, max_rows=200)
    assert (out["total_rows"], out["returned_rows"], out["truncated"]) == (450, 50, False)
    assert len(fake_origin.get_worksheet_calls) == 1


def test_start_row_past_the_end_raises_with_the_row_count(fake_origin):
    fake_origin.worksheet_data = _rows(10)
    with pytest.raises(ValueError, match=r"past the end .*\(10 rows\)"):
        _read(start_row=11)


def test_existing_empty_sheet_reads_as_zero_rows(fake_origin):
    # Origin 2018 answers GetWorksheet on an empty sheet with an HRESULT int,
    # full read and range read alike; the sheet exists, so it is not "missing".
    fake_origin.GetWorksheet = lambda target, *rng: HRESULT
    out = _read()
    assert (out["total_rows"], out["total_cols"], out["columns"]) == (0, 0, [])
    with pytest.raises(ValueError, match="past the end"):
        _read(start_row=1)


def test_missing_sheet_is_still_not_found(fake_origin):
    with pytest.raises(ValueError, match="not found"):
        json.loads(get_worksheet_data("Ghost", "Sheet1"))


def test_build_without_range_arguments_falls_back_to_a_full_read(fake_origin):
    data = _rows(450)

    def full_only(target, *rng):
        if rng:
            raise TypeError("GetWorksheet() takes 1 positional argument")
        return data

    fake_origin.GetWorksheet = full_only
    out = _read(start_row=200, max_rows=200)
    assert out["columns"][0] == [float(i) for i in range(200, 400)]
    assert (out["total_rows"], out["next_start_row"]) == (450, 400)


def test_max_rows_zero_still_reads_everything_in_one_call(fake_origin):
    fake_origin.worksheet_data = _rows(450)
    out = _read(max_rows=0)
    assert out["returned_rows"] == 450 and out["truncated"] is False
    assert [c[1] for c in fake_origin.get_worksheet_calls] == [None]
