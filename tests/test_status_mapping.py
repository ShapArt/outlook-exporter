import tempfile

import pandas as pd

from config import AppConfig, Paths
from core import db, excel, sla
from core.sla import STATUS_MODEL, status_code_to_label, status_text_to_code


def _temp_cfg():
    tmp = tempfile.mkdtemp()
    paths = Paths(
        appdata_dir=tmp,
        log_dir=f"{tmp}/logs",
        db_path=f"{tmp}/test.sqlite3",
        excel_path=f"{tmp}/tickets.xlsx",
        backup_dir=f"{tmp}/backups",
    )
    return AppConfig(paths=paths)


def test_status_mapping_ru():
    for code, label in STATUS_MODEL:
        assert status_text_to_code(label) == code
        assert status_code_to_label(code) == label


def test_status_mapping_aliases():
    assert status_text_to_code("resolved") == sla.STATUS_RESOLVED
    assert status_text_to_code("waiting customer") == sla.STATUS_WAITING_CUSTOMER


def test_export_contains_ru_labels():
    cfg = _temp_cfg()
    db.ensure_schema(cfg)
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    ticket_id = db.seed_test_ticket(conn, status=sla.STATUS_ASSIGNED)
    conn.close()

    path = excel.export_excel(cfg, today_only=False)
    df = pd.read_excel(path, sheet_name=excel.SHEET_TICKETS)
    assert not df.empty
    allowed_labels = {label for _, label in STATUS_MODEL}
    assert set(df["Статус"].unique()) <= allowed_labels
    # sanity: the seeded ticket id is present
    assert ticket_id in set(df["ticket_id"])


def test_sync_accepts_ru_labels():
    cfg = _temp_cfg()
    db.ensure_schema(cfg)
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    ticket_id = db.seed_test_ticket(conn, status=sla.STATUS_NEW)
    conn.close()

    path = excel.export_excel(cfg, today_only=False)
    df = pd.read_excel(path, sheet_name=excel.SHEET_TICKETS)
    df.loc[df["ticket_id"] == ticket_id, "Статус"] = status_code_to_label(
        sla.STATUS_RESOLVED
    )
    df.to_excel(path, sheet_name=excel.SHEET_TICKETS, index=False)

    res = excel.sync_from_excel(cfg)
    assert res["updated"] >= 1

    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    row = conn.execute(
        "SELECT status, row_version FROM tickets WHERE id=?", (ticket_id,)
    ).fetchone()
    conn.close()
    assert row["status"] == sla.STATUS_RESOLVED
    assert int(row["row_version"]) >= 2
