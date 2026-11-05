"""Desktop UI tests (offscreen): workflow end-to-end and event-loop responsiveness."""

from __future__ import annotations

import shutil
import time

import pytest

pytest.importorskip("PySide6.QtWidgets")
pytestmark = pytest.mark.ui

from PySide6.QtCore import Qt, QTimer  # noqa: E402

from docgen.i18n.catalog import set_language, tr  # noqa: E402
from docgen.testing.samples import write_customer_a  # noqa: E402
from docgen.ui.main_window import MainWindow  # noqa: E402
from docgen.ui.state import AppState  # noqa: E402


@pytest.fixture(autouse=True)
def no_modal_dialogs(monkeypatch):
    """Record message boxes instead of blocking the test on a modal dialog."""
    from PySide6.QtWidgets import QMessageBox

    shown: list[str] = []

    def record(*args, **kwargs):
        shown.append(" | ".join(str(a) for a in args[1:3]))
        return QMessageBox.StandardButton.Yes

    def record_exec(self):
        shown.append(f"{self.text()} | {self.detailedText()[:500]}")
        return 0

    monkeypatch.setattr(QMessageBox, "information", staticmethod(record))
    monkeypatch.setattr(QMessageBox, "question", staticmethod(record))
    monkeypatch.setattr(QMessageBox, "exec", record_exec)
    return shown


@pytest.fixture(autouse=True)
def gui_thread_gc(qapp):
    from docgen.ui.gc_guard import MainThreadGarbageCollector

    collector = MainThreadGarbageCollector(qapp)
    yield collector
    collector.stop()


@pytest.fixture()
def window(qtbot, tmp_path):
    set_language("en")
    state = AppState(tmp_path / "home")
    profile, workbook = write_customer_a(tmp_path / "src", n_employees=30)
    state.repo.import_bundle(profile)
    shutil.copytree(
        profile.parent / "templates", state.repo.directory("customer_a") / "templates", dirs_exist_ok=True
    )
    state.open_profile("customer_a")
    win = MainWindow(state)
    qtbot.addWidget(win)
    win.show()
    win.pages["excel"].set_excel(workbook)
    qtbot.waitUntil(lambda: state.workbook is not None, timeout=20000)
    return win, state, tmp_path


def test_navigation_and_pages_render(window, qtbot):
    win, state, _ = window
    for key in (
        "dashboard",
        "profiles",
        "excel",
        "relationships",
        "mapping",
        "templates",
        "preview",
        "generate",
        "logs",
        "settings",
    ):
        win.go(key)
        qtbot.wait(10)
        assert win.stack.currentWidget().isVisible()
    assert win.pages["excel"].sheets.count() == 5
    assert win.pages["relationships"].list.count() == 4
    assert win.pages["templates"].list.count() == 3


def test_mapping_designer_edits_and_tests_a_field(window, qtbot, no_modal_dialogs):
    win, state, _ = window
    page = win.pages["mapping"]
    win.go("mapping")
    page._add_field()
    editor = page.field_editor
    editor.name.setText("employee.dept_upper")
    editor.kind.setCurrentIndex(1)  # lookup chain
    editor.chain._add_step("emp_dept")
    editor.chain.column.setCurrentText("name")
    assert (
        "Employees.DepartmentID → Departments.DepartmentID → Departments.DepartmentName"
        in editor.chain.breadcrumb.text()
    )
    editor.transforms._append("upper", {})
    page._apply()
    assert not no_modal_dialogs, no_modal_dialogs
    assert state.dirty
    rule = state.profile.fields[-1]
    assert rule.path == "emp_dept.name" and rule.transforms[0].name == "upper"
    page.test_row.setValue(4)
    page._test()
    qtbot.waitUntil(
        lambda: any(
            "employee.dept_upper = 'PHÒNG KẾ TOÁN'" in page.test_output.item(i).text()
            for i in range(page.test_output.count())
        ),
        timeout=20000,
    )
    assert not no_modal_dialogs, no_modal_dialogs


def test_generate_from_ui_stays_responsive(window, qtbot):
    win, state, tmp_path = window
    page = win.pages["generate"]
    win.go("generate")
    page.output.set_path(tmp_path / "out")
    page.workers.setValue(1)
    gaps: list[float] = []
    last = [time.perf_counter()]

    def beat() -> None:
        now = time.perf_counter()
        gaps.append(now - last[0])
        last[0] = now

    timer = QTimer()
    timer.timeout.connect(beat)
    timer.start(20)
    page._start(dry_run=False)
    qtbot.waitUntil(lambda: page.worker is None and page.result is not None, timeout=120000)
    timer.stop()
    assert page.result.status == "completed"
    assert len(list((tmp_path / "out").glob("*.docx"))) == 30
    assert page.units.rowCount() == 30
    # UI thread kept processing events while 30 merged documents were generated in the background
    worst = max(gaps[1:]) if len(gaps) > 1 else 0
    assert len(gaps) > 10
    assert worst < 0.5, f"event loop stalled for {worst:.3f}s"
    jobs = state.db.jobs()
    assert jobs and jobs[0].documents == 30


def test_cancel_from_ui(window, qtbot):
    win, state, tmp_path = window
    page = win.pages["generate"]
    page.output.set_path(tmp_path / "out2")
    page.workers.setValue(1)
    page._start(dry_run=False)
    qtbot.waitUntil(lambda: page.progress.maximum() > 0 and page.progress.value() >= 1, timeout=60000)
    page._cancel()
    qtbot.waitUntil(lambda: page.worker is None, timeout=60000)
    assert page.result.status == "cancelled"
    assert not list((tmp_path / "out2").rglob("*.tmp"))


def test_vietnamese_catalog_is_complete_for_ui():
    set_language("vi")
    try:
        assert tr("Generate") == "Tạo tài liệu"
        assert tr("{errors} error(s), {warnings} warning(s)", errors=1, warnings=2) == "1 lỗi, 2 cảnh báo"
    finally:
        set_language("en")


def test_templates_page_lists_form_inputs_and_binding_fills_them(qtbot, tmp_path, no_modal_dialogs):
    from docgen.testing.samples import write_customer_c

    set_language("en")
    state = AppState(tmp_path / "home")
    profile, workbook = write_customer_c(tmp_path / "src")
    state.repo.import_bundle(profile)
    shutil.copytree(
        profile.parent / "templates", state.repo.directory("customer_c") / "templates", dirs_exist_ok=True
    )
    state.open_profile("customer_c")
    win = MainWindow(state)
    qtbot.addWidget(win)
    win.show()
    page = win.pages["templates"]
    win.go("templates")
    table = page.placeholders

    def rows() -> dict[str, tuple[str, str]]:
        return {
            table.item(r, 0).text(): (table.item(r, 1).text(), table.item(r, 2).text())
            for r in range(table.rowCount())
        }

    qtbot.waitUntil(lambda: "Người khai ký tên" in rows(), timeout=20000)
    current = rows()
    assert current["Họ và tên"] == ("blank", "will be filled")
    assert current["phong_ban"] == ("content control", "will be filled")
    assert current["MST"] == ("form field", "will be filled")
    assert current["Người khai ký tên"] == ("blank", "left as is")
    # bind the signature date blank to a field -> after the debounced re-scan it will be filled
    row = next(r for r in range(table.rowCount()) if table.item(r, 0).text() == "ngày")
    combo = table.cellWidget(row, 3)
    combo.setCurrentIndex(combo.findData("nhan_vien.sinh_ngay"))
    assert state.profile.templates[0].bindings["ngày"] == "nhan_vien.sinh_ngay"
    qtbot.waitUntil(lambda: rows().get("ngày", ("", ""))[1] == "will be filled", timeout=20000)
    assert not no_modal_dialogs


def test_quick_start_wizard_from_excel_and_word_to_documents(qtbot, tmp_path):
    """A first-time user: no profile, just an Excel file and two Word forms."""
    from docgen.testing.construction import write_construction_kit

    set_language("vi")
    state = AppState(tmp_path / "home")
    _, workbook = write_construction_kit(tmp_path / "files")
    win = MainWindow(state)
    qtbot.addWidget(win)
    win.show()
    win.go("quickstart")
    page = win.pages["quickstart"]
    assert not page.next_btn.isEnabled()
    page.open_excel(workbook)
    qtbot.waitUntil(lambda: page.analysis is not None and page.worker is None, timeout=30000)
    assert "6 dòng – mỗi dòng sẽ tạo một bộ tài liệu" in page.found.text()
    assert page.next_btn.isEnabled()
    page.next()
    templates = tmp_path / "files" / "templates"
    page.add_words([templates / "PhieuYeuCauNghiemThu.docx", templates / "BienBanNghiemThu.docx"])
    page.next()
    qtbot.waitUntil(lambda: page.stack.currentIndex() == 2 and page.worker is None, timeout=30000)
    assert "36" in page.match_info.text()  # 36 of 38 places matched automatically
    page.choose(0, "Ngày yêu cầu bằng chữ", "dong.ngay_nghiem_thu_bang_chu")
    page.next()
    assert page.name.text() == "Nhà làm việc 5 tầng – Công ty CP Đầu tư DEF"
    page.output.edit.setText(str(tmp_path / "ket qua"))
    page.next()  # save and create
    generate = win.pages["generate"]
    qtbot.waitUntil(lambda: generate.result is not None and generate.worker is None, timeout=120000)
    assert generate.result.status == "completed", [str(i) for i in generate.result.all_issues()]
    assert len(generate.result.documents) == 6 and generate.result.combined_path
    assert state.profile is not None and state.repo.exists(state.profile.profile.id)
    assert win.nav.currentItem().data(Qt.ItemDataRole.UserRole) == "generate"


def test_simple_mode_hides_setup_pages_and_example_runs(qtbot, tmp_path, monkeypatch):
    from docgen.ui.main_window import ADVANCED_PAGES

    monkeypatch.setattr(
        "docgen.ui.pages.dashboard.documents_folder", lambda: tmp_path / "Documents" / "DocGen"
    )
    set_language("vi")
    state = AppState(tmp_path / "home")
    win = MainWindow(state)
    qtbot.addWidget(win)
    win.show()
    visible = {
        win.nav.item(i).data(Qt.ItemDataRole.UserRole)
        for i in range(win.nav.count())
        if not win.nav.item(i).isHidden()
    }
    assert visible == {"dashboard", "quickstart", "profiles", "generate", "logs", "settings"}
    state.settings.advanced_mode = True
    state.settings_changed.emit()
    assert all(not win.nav.item(i).isHidden() for i in range(win.nav.count()))
    state.settings.advanced_mode = False
    state.settings_changed.emit()
    assert {
        win.nav.item(i).data(Qt.ItemDataRole.UserRole)
        for i in range(win.nav.count())
        if win.nav.item(i).isHidden()
    } >= ADVANCED_PAGES

    home = win.pages["dashboard"]
    home.try_sample()  # "Dùng thử hồ sơ mẫu"
    folder = tmp_path / "Documents" / "DocGen" / "Mẫu - Hồ sơ nghiệm thu"
    assert (folder / "data" / "HoSoNghiemThu.xlsx").exists() and (folder / "templates").is_dir()
    generate = win.pages["generate"]
    qtbot.waitUntil(lambda: generate.result is not None and generate.worker is None, timeout=120000)
    assert generate.result.status == "completed", [str(i) for i in generate.result.all_issues()]
    assert len(list((folder / "Kết quả").glob("*.docx"))) == 7  # 6 work items + the combined file
    assert "Tổng hợp hồ sơ nghiệm thu.docx" in generate.stage.text()
    # the user's edits to the example Excel file survive a second click
    excel = folder / "data" / "HoSoNghiemThu.xlsx"
    excel.write_bytes(excel.read_bytes() + b"")
    stamp = excel.stat().st_mtime_ns
    from docgen.application.starter import install_construction_kit

    install_construction_kit(state.repo, folder)
    assert excel.stat().st_mtime_ns == stamp
    assert (
        home.profiles.count() == 1
        and home.profiles.item(0).text() == "Hồ sơ nghiệm thu công việc xây dựng (mẫu)"
    )
