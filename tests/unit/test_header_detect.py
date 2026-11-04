from docgen.infrastructure.excel.header_detect import combine_header_rows, detect_header_row


def test_detects_header_below_title_rows():
    rows = [
        ["BÁO CÁO"],
        [],
        ["Ngày: 01/01/2024", None],
        ["Mã", "Họ tên", "Phòng"],
        ["E1", "An", "IT"],
        ["E2", "Bình", "KT"],
    ]
    result = detect_header_row(rows)
    assert result.row == 4


def test_expected_headers_win():
    rows = [["a", "b", "c", "d"], ["EmployeeID", "FullName", None, None], ["E1", "An", None, None]]
    assert detect_header_row(rows, expected=["EmployeeID", "FullName"]).row == 2
    assert detect_header_row(rows, expected=["employeeid"], casefold=True).confidence == 1.0


def test_accent_insensitive_matching():
    rows = [["Ma NV", "Ho ten"], ["1", "x"]]
    assert (
        detect_header_row(rows, expected=["Mã NV", "Họ tên"], accents=True).candidates[0].expected_hits == 2
    )


def test_no_header():
    assert detect_header_row([[1, 2], [3, 4]]).row is None


def test_combine_multi_row_headers_with_merged_cells():
    rows = [["Thông tin", None, "Lương", None], ["Mã", "Họ tên", "Cơ bản", "Phụ cấp"]]
    assert combine_header_rows(rows) == ["Thông tin Mã", "Thông tin Họ tên", "Lương Cơ bản", "Lương Phụ cấp"]
