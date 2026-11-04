import datetime as dt
import zipfile

import pytest

from docgen.domain.naming import render_pattern, sanitize_component, sanitize_relative_path
from docgen.domain.profile.models import ConflictPolicy
from docgen.infrastructure.fs.atomic import OutputConflictError, finalize, temp_path_for
from docgen.infrastructure.fs.safety import UnsafeFileError, ZipLimits, check_zip, is_within


def test_render_pattern():
    values = {"e": {"code": "E1", "dept": "IT/Kế toán", "date": dt.date(2024, 5, 6), "n": 7}}
    text, missing = render_pattern("HS_{e.code}_{e.dept}_{e.date:%d%m%Y}_{e.n:03d}_{e.none}{{x}}", values)
    assert text == "HS_E1_IT-Kế toán_06052024_007_{x}" and missing == ["e.none"]


@pytest.mark.parametrize(
    ("name", "expected"),
    [
        ("a<b>:c?.docx", "a_b__c_.docx"),
        ("CON.docx", "_CON.docx"),
        ("lpt1", "_lpt1"),
        ("  x. ", "x"),
        ("", "untitled"),
        ("Hồ sơ.docx", "Hồ sơ.docx"),
    ],
)
def test_sanitize_component(name, expected):
    assert sanitize_component(name) == expected


def test_long_names_are_shortened_with_hash():
    name = "N" * 300 + ".docx"
    short = sanitize_component(name, max_length=80)
    assert len(short) <= 80 and short.endswith(".docx") and "~" in short
    assert sanitize_relative_path("../a//b\\c/../d.docx") == "a/b/c/d.docx"


def test_finalize_policies(tmp_path):
    target = tmp_path / "out" / "r.docx"
    for _ in range(2):
        tmp = temp_path_for(target)
        tmp.write_text("x")
        res = finalize(tmp, target, ConflictPolicy.SUFFIX)
    assert res.path == tmp_path / "out" / "r (2).docx" and res.renamed_from == target
    tmp = temp_path_for(target)
    tmp.write_text("y")
    assert finalize(tmp, target, ConflictPolicy.SKIP).skipped and not tmp.exists()
    tmp = temp_path_for(target)
    tmp.write_text("z")
    with pytest.raises(OutputConflictError):
        finalize(tmp, target, ConflictPolicy.ERROR)
    tmp = temp_path_for(target)
    tmp.write_text("new")
    assert finalize(tmp, target, ConflictPolicy.OVERWRITE).path == target and target.read_text() == "new"
    reserved = {tmp_path / "out" / "same.docx"}
    tmp = temp_path_for(target)
    tmp.write_text("q")
    res = finalize(tmp, tmp_path / "out" / "same.docx", ConflictPolicy.OVERWRITE, reserved=reserved)
    assert res.path.name == "same (2).docx"
    assert not list((tmp_path / "out").glob("*.tmp"))


def test_zip_safety(tmp_path):
    bomb = tmp_path / "bomb.xlsx"
    with zipfile.ZipFile(bomb, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("xl/big.xml", b"0" * (30 * 1024 * 1024))
    with pytest.raises(UnsafeFileError, match="ratio"):
        check_zip(bomb)
    with pytest.raises(UnsafeFileError, match="uncompressed"):
        check_zip(bomb, ZipLimits(max_total_uncompressed=1024))
    notzip = tmp_path / "x.docx"
    notzip.write_text("hello")
    with pytest.raises(UnsafeFileError):
        check_zip(notzip)
    assert is_within(tmp_path / "a" / "b", tmp_path) and not is_within(tmp_path / ".." / "x", tmp_path)
