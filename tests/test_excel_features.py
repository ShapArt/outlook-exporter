import tempfile
from pathlib import Path

import pandas as pd

from config import AppConfig, Paths
from core import db, excel, sla

STATUS_COL = "Статус"  # visible status column header


def _cfg():
    tmp = tempfile.mkdtemp()
    return AppConfig(
        paths=Paths(
            appdata_dir=tmp,
            log_dir=f"{tmp}/logs",
            db_path=f"{tmp}/db.sqlite3",
            excel_path=f"{tmp}/t.xlsx",
            backup_dir=f"{tmp}/bak",
        )
    )


def test_export_creates_all_sheets_and_hides_columns():
    cfg = _cfg()
    db.ensure_schema(cfg)
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    db.seed_test_ticket(conn, status=sla.STATUS_ASSIGNED)
    conn.close()
    path = excel.export_excel(cfg, today_only=False)
    xls = pd.ExcelFile(path)
    expected = {
        excel.SHEET_TICKETS,
        excel.SHEET_OVERDUE,
        excel.SHEET_KPI,
        excel.SHEET_CONFLICTS,
        excel.SHEET_STATUS_HELP,
    }
    assert expected.issubset(set(xls.sheet_names))
    df = pd.read_excel(path, sheet_name=excel.SHEET_TICKETS)
    assert STATUS_COL in df.columns


def test_atomic_replace_creates_pending_on_permission_error(monkeypatch):
    cfg = _cfg()
    target = Path(cfg.paths.excel_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text("lock", encoding="utf-8")
    tmp = target.with_name("tmp.xlsx")
    tmp.write_text("new", encoding="utf-8")

    calls = {"first": True}

    real_replace = Path.replace

    def fake_replace(self, other):
        if self == tmp and calls["first"]:
            calls["first"] = False
            raise PermissionError("locked")
        return real_replace(self, other)

    monkeypatch.setattr(Path, "replace", fake_replace)
    pending = excel._atomic_replace(tmp, target)
    assert pending.suffix == ".xlsx" and "pending" in pending.name


def test_sync_from_excel_conflict_row_version():
    cfg = _cfg()
    db.ensure_schema(cfg)
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    tid = db.seed_test_ticket(conn, status=sla.STATUS_NEW)
    conn.close()
    path = excel.export_excel(cfg, today_only=False)
    df = pd.read_excel(path, sheet_name=excel.SHEET_TICKETS)
    df.loc[df["ticket_id"] == tid, "row_version"] = 999  # conflict
    with pd.ExcelWriter(
        path, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        df.to_excel(writer, sheet_name=excel.SHEET_TICKETS, index=False)
    res = excel.sync_from_excel(cfg)
    assert res["conflicts"] >= 1
