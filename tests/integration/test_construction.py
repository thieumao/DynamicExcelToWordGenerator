"""Construction starter kit: one ready-to-print file per work item, values from three sheets."""

from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

from docgen.application.generation.job import GenerationRequest
from docgen.application.generation.service import DocumentGenerationService
from docgen.infrastructure.word.verifier import verify_docx

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def text_of(path: Path) -> str:
    root = etree.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    return "\n".join("".join(t.text or "" for t in p.iter(W + "t")) for p in root.iter(W + "p"))


def test_one_file_per_work_item_with_request_and_record(construction, tmp_path):
    result = DocumentGenerationService().run(
        GenerationRequest(
            construction.profile, construction.directory, construction.workbook, tmp_path / "out"
        )
    )
    assert result.status == "completed", [str(i) for i in result.all_issues()]
    names = sorted(Path(p).name for p in result.documents)
    assert names[0] == "1 - Lắp dựng cốt thép cột tầng 1.docx" and len(names) == 6
    text = text_of(tmp_path / "out" / names[0])
    for expected in (
        "Hà Nội, ngày 01 tháng 04 năm 2024",  # request: the day before the acceptance
        "PHIẾU YÊU CẦU NGHIỆM THU",
        "Kính gửi: Công ty TNHH Tư vấn Xây dựng XYZ",  # from the general information sheet
        "Hà Nội, ngày 02 tháng 04 năm 2024",
        "BIÊN BẢN NGHIỆM THU CÔNG VIỆC XÂY DỰNG",
        "Số: 01/NTCV",
        "2. Đối tượng nghiệm thu: Lắp dựng cốt thép cột tầng 1",
        "Ông/Bà: Nguyễn Văn Hùng – Chức vụ: Kỹ sư tư vấn giám sát",  # looked up by code GS01
        "Ông/Bà: Trần Minh Tuấn – Chức vụ: Cán bộ kỹ thuật thi công",
        "Bắt đầu: 08 giờ 00 ngày 02 tháng 04 năm 2024",
        "Kết thúc: 09 giờ 30 ngày 02 tháng 04 năm 2024",
        "– Tiêu chuẩn áp dụng: TCVN 4453:1995; TCVN 1651-2:2018.",
        "b) Về chất lượng công việc xây dựng: Đạt yêu cầu",
    ):
        assert expected in text, expected
    assert "c) Các ý kiến khác (nếu có): ......" in text  # unbound blank stays for handwriting
    assert "<<" not in text and "{{" not in text
    assert "Phạm Quốc Bảo" in text_of(tmp_path / "out" / names[2])  # item 3: engineer KT02


def test_everything_in_one_file_for_printing(construction, tmp_path):
    result = DocumentGenerationService().run(
        GenerationRequest(
            construction.profile, construction.directory, construction.workbook, tmp_path / "out"
        )
    )
    assert result.status == "completed", [str(i) for i in result.all_issues()]
    combined = Path(result.combined_path)
    assert combined.name == "Tổng hợp hồ sơ nghiệm thu.docx" and combined.parent == tmp_path / "out"
    text = text_of(combined)
    assert text.count("BIÊN BẢN NGHIỆM THU CÔNG VIỆC XÂY DỰNG") == 6
    assert text.count("PHIẾU YÊU CẦU NGHIỆM THU") == 6
    assert text.index("Lắp dựng cốt thép cột tầng 1") < text.index("Đổ bê tông dầm, sàn tầng 2")  # in order
    assert verify_docx(combined).ok
    # the user can switch it off for one run
    again = DocumentGenerationService().run(
        GenerationRequest(
            construction.profile,
            construction.directory,
            construction.workbook,
            tmp_path / "o2",
            combined=False,
        )
    )
    assert again.combined_path == "" and len(again.documents) == 6
