import json

from docgen.cli.main import main, parse_rows


def test_parse_rows():
    assert parse_rows("5,7-9") == [5, 7, 8, 9] and parse_rows(None) is None


def test_cli_end_to_end(tmp_path, capsys):
    assert main(["samples", str(tmp_path / "s"), "--employees", "3"]) == 0
    a = tmp_path / "s" / "customer_a"
    assert main(["inspect", str(a / "data" / "HR.xlsx")]) == 0
    assert "header row: 3" in capsys.readouterr().out
    assert main(["profile", "validate", str(a / "profile.yaml"), "--excel", str(a / "data" / "HR.xlsx")]) == 0
    assert main(["templates", "scan", str(a / "templates" / "EmployeeDetails.docx")]) == 0
    capsys.readouterr()
    assert (
        main(
            [
                "preview",
                "--profile",
                str(a / "profile.yaml"),
                "--excel",
                str(a / "data" / "HR.xlsx"),
                "--rows",
                "4",
                "--json",
            ]
        )
        == 0
    )
    payload = json.loads(capsys.readouterr().out)
    assert payload[0]["context"]["employee"]["code"] == "E001"
    out = tmp_path / "out"
    assert (
        main(
            [
                "generate",
                "--profile",
                str(a / "profile.yaml"),
                "--excel",
                str(a / "data" / "HR.xlsx"),
                "--out",
                str(out),
                "--quiet",
            ]
        )
        == 0
    )
    assert len(list(out.glob("*.docx"))) == 3
    schema = tmp_path / "schema.json"
    assert main(["profile", "schema", "-o", str(schema)]) == 0
    assert "relationships" in json.loads(schema.read_text(encoding="utf-8"))["properties"]


def test_cli_reports_invalid_profile(tmp_path, capsys):
    bad = tmp_path / "bad.yaml"
    bad.write_text("profile: {id: x}\n", encoding="utf-8")
    assert main(["profile", "validate", str(bad)]) == 1
    assert "INVALID" in capsys.readouterr().out
