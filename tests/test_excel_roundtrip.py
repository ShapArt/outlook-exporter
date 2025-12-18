import tempfile

import pandas as pd

from config import AppConfig, Paths
from core import db, excel, sla

STATUS_COL = "Статус"  # status column header
RESP_COL = "Ответственный"  # responsible column header
PRIO_COL = "Приоритет"  # priority column header


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


def test_excel_roundtrip_updates_db_and_export():
    cfg = _temp_cfg()
    db.ensure_schema(cfg)
    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    ticket_new = db.seed_test_ticket(conn, status=sla.STATUS_NEW)
    ticket_assigned = db.seed_test_ticket(conn, status=sla.STATUS_ASSIGNED)
    conn.close()

    path = excel.export_excel(cfg, today_only=False)

    df = pd.read_excel(path, sheet_name=excel.SHEET_TICKETS)
    df.loc[df["ticket_id"] == ticket_new, STATUS_COL] = sla.status_code_to_label(
        sla.STATUS_RESOLVED
    )
    df.loc[df["ticket_id"] == ticket_new, RESP_COL] = "user1@example.com"
    df.loc[df["ticket_id"] == ticket_new, PRIO_COL] = "p2"
    df.loc[df["ticket_id"] == ticket_assigned, STATUS_COL] = sla.status_code_to_label(
        sla.STATUS_WAITING_CUSTOMER
    )
    with pd.ExcelWriter(
        path, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        df.to_excel(writer, sheet_name=excel.SHEET_TICKETS, index=False)

    res = excel.sync_from_excel(cfg)
    assert res["updated"] >= 1

    conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
    row1 = conn.execute(
        "SELECT status, responsible, priority FROM tickets WHERE id=?", (ticket_new,)
    ).fetchone()
    row2 = conn.execute(
        "SELECT status FROM tickets WHERE id=?", (ticket_assigned,)
    ).fetchone()
    conn.close()

    assert row1["status"] == sla.STATUS_RESOLVED
    assert row1["responsible"] == "user1@example.com"
    assert row1["priority"] == "p2"
    assert row2["status"] == sla.STATUS_WAITING_CUSTOMER

    path2 = excel.export_excel(cfg, today_only=False)
    df2 = pd.read_excel(path2, sheet_name=excel.SHEET_TICKETS)
    status_exported = df2.loc[df2["ticket_id"] == ticket_new, STATUS_COL].iloc[0]
    assert status_exported == sla.status_code_to_label(sla.STATUS_RESOLVED)
