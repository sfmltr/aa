"""Core logic tests (no GUI). Run: python -m pytest tests -q"""
from pathlib import Path

import openpyxl
import pytest
from openpyxl.styles import PatternFill

from ncr_panel.config import AppConfig
from ncr_panel.excel_data import extract_part_numbers, categorize, load_excel
from ncr_panel.file_search import FileIndex
from ncr_panel.repetitive import match_record, scan_notes
from ncr_panel.utils import normalize_text, col_letter_to_index


def _workbook(path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["info"])
    ws.append(["Note", "Date", "MRB ENG.", "NCR NO", "P", "R", "Discrepancy", "Status", "X", "Part No",
               "K", "L", "M", "N", "O", "P2", "CN/CSI"])
    ws.append(["", "d", "Ali", "NCR-1", "", "", "Crack on 123A4567-1 rib", "MRB-ENG", "", "401B2222-5",
               "", "", "", "", "", "", "CN"])
    ws.append(["Stress check", "d", "Ali", "NCR-2", "", "", "Hole on 234B5678-12", "EXTERNAL APPROVER", "",
               "401B2222-5", "", "", "", "", "", "", "CSI"])
    ws.append(["", "d", "Ayse", "NCR-3", "", "", "Scratch 345C6789", "CLOSED", "", "401B2222-5",
               "", "", "", "", "", "", "CN"])
    for c in range(1, 18):
        ws.cell(3, c).fill = PatternFill(fill_type="solid", fgColor="FFFFC7CE")
    wb.save(path)


@pytest.fixture
def cfg(tmp_path: Path) -> AppConfig:
    cfg = AppConfig()
    xl = tmp_path / "ncr.xlsx"
    _workbook(xl)
    cfg.excel.path = str(xl)
    cfg.excel.color_source = "openpyxl"
    return cfg


def test_col_letters():
    assert col_letter_to_index("A") == 1
    assert col_letter_to_index("Q") == 17
    assert col_letter_to_index("AA") == 27


def test_part_numbers():
    assert extract_part_numbers("see 123A4567-1 and 123a4567-1 and 999Z9999", r"\b\d{3}[A-Z]\d{4}(?:-\d+)*\b") == \
        ["123A4567-1", "999Z9999"]
    assert extract_part_numbers("456D7890-3_report", r"\b\d{3}[A-Z]\d{4}(?:-\d+)*\b") == ["456D7890-3"]


def test_categorize():
    cats = {"Stress": ["stress"], "SNN": ["snn"]}
    assert categorize("Needs STRESS and snn", cats) == ["Stress", "SNN"]
    assert categorize("", cats) == []


def test_load_excel(cfg: AppConfig):
    data = load_excel(cfg)
    assert data.header_row == 2
    assert [r.ncr_no for r in data.records] == ["NCR-1", "NCR-2", "NCR-3"]
    assert data.engineers == ["Ali", "Ayse"]
    kinds = [r.kind(cfg.excel) for r in data.records]
    assert kinds == ["mrb", "external", "other"]
    assert data.records[0].color == "#FFC7CE"
    assert data.records[1].color is None
    assert data.records[1].categories == ["Stress"]
    assert data.records[0].part_numbers == ["123A4567-1"]
    assert data.records[0].extra["Part No"] == "401B2222-5"


def test_file_index_matches_by_discrepancy_not_part_no(cfg: AppConfig, tmp_path: Path):
    folder = tmp_path / "docs"
    folder.mkdir()
    (folder / "a.txt").write_text("repair of 123A4567-1 approved")
    (folder / "b.txt").write_text("assembly 401B2222-5 only")          # the Part No column -> must NOT match
    (folder / "c.txt").write_text("mentions NCR-1 without part")
    cfg.files.folders = [str(folder)]
    data = load_excel(cfg)
    idx = FileIndex()
    assert idx.rebuild(cfg.files) == 3
    hits = idx.match_record(data.records[0], cfg.files)
    names = {Path(h.entry.path).name: h for h in hits}
    assert set(names) == {"a.txt", "c.txt"}
    assert names["a.txt"].by_part and not names["c.txt"].by_part
    idx.annotate_records(data.records, cfg.files)
    assert data.records[0].file_match_count == 1
    assert data.records[0].ncr_file_match_count == 1
    # incremental rebuild reads nothing new
    assert idx.rebuild(cfg.files) == 0


def test_repetitive(cfg: AppConfig, tmp_path: Path):
    root = tmp_path / "rep"
    (root / "Bracket").mkdir(parents=True)
    (root / "Bracket" / "note.txt").write_text("Crack on 123A4567-1 rib\nrepair per SRM")
    (root / "Other").mkdir()
    (root / "Other" / "note.txt").write_text("completely different issue")
    cfg.repetitive.folders = [str(root)]
    entries = scan_notes(cfg.repetitive)
    assert len(entries) == 2
    data = load_excel(cfg)
    hits = match_record(data.records[0], entries, cfg.repetitive)
    assert len(hits) == 1 and hits[0].exact and hits[0].entry.name == "Bracket"
    assert match_record(data.records[1], entries, cfg.repetitive) == []


def test_normalize():
    assert normalize_text("  Çatlak,  Bracket!  ") == "catlak bracket"


def test_config_roundtrip(tmp_path: Path):
    cfg = AppConfig()
    cfg.excel.path = "x.xlsm"
    cfg.ui.quick_folders = ["a", "b"]
    p = tmp_path / "c.json"
    cfg.save(p)
    loaded = AppConfig.load(p)
    assert loaded.excel.path == "x.xlsm" and loaded.ui.quick_folders == ["a", "b"]
    assert loaded.outlook.days_back == cfg.outlook.days_back
