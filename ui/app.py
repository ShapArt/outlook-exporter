from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Callable, Dict, List, Optional

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices

from config import AppConfig, default_config
from core import db, excel, notify, sla
from core.logger import get_logger, register_ui_sink, setup_logging
from core.outlook import OutlookClient, detect_outlook_environment

log = get_logger(__name__)
NARROW_THRESHOLD = 1200


class TaskThread(QtCore.QThread):
    finished_ok = QtCore.Signal(str)
    finished_err = QtCore.Signal(str)
    result_ready = QtCore.Signal(object)

    def __init__(self, label: str, fn: Callable[[], object]):
        super().__init__()
        self.label = label
        self.fn = fn

    def run(self):
        try:
            res = self.fn()
            self.result_ready.emit(res)
            suffix = f": {res}" if res not in (None, "", {}) else ""
            self.finished_ok.emit(f"{self.label}: готово{suffix}")
        except Exception as exc:
            err = f"{self.label}: ошибка {exc}"
            log.exception(err)
            self.finished_err.emit(err)


class QtLogBridge(QtCore.QObject):
    log_received = QtCore.Signal(str)

    def __init__(self, target: Callable[[str], None]):
        super().__init__()
        self.log_received.connect(target)
        register_ui_sink(self.log_received.emit)


class LogConsole(QtWidgets.QWidget):
    command_entered = QtCore.Signal(str)

    def __init__(self):
        super().__init__()
        layout = QtWidgets.QVBoxLayout()
        self.view = QtWidgets.QPlainTextEdit()
        self.view.setReadOnly(True)
        self.view.setMaximumBlockCount(2000)
        self.view.setStyleSheet(
            "font-family: Consolas, 'Cascadia Code', monospace; background:#0f172a; color:#e2e8f0;"
        )
        self.input = QtWidgets.QLineEdit()
        self.input.setPlaceholderText(
            "Команды: help / sync / ingest / export / send / process / diagnose / open excel"
        )
        self.input.returnPressed.connect(self._on_enter)
        layout.addWidget(self.view, 1)
        layout.addWidget(self.input, 0)
        self.setLayout(layout)

    def append_line(self, text: str):
        self.view.appendPlainText(text)
        self.view.moveCursor(QtGui.QTextCursor.End)

    def _on_enter(self):
        cmd = self.input.text().strip()
        if not cmd:
            return
        self.append_line(f"> {cmd}")
        self.input.clear()
        self.command_entered.emit(cmd)


class MetricCard(QtWidgets.QFrame):
    def __init__(self, title: str):
        super().__init__()
        self.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.setStyleSheet(
            """
            QFrame {border:1px solid #e2e8f0; border-radius:10px; background:#f8fafc;}
            QLabel.title {color:#475569; font-weight:600;}
            """
        )
        layout = QtWidgets.QVBoxLayout()
        self.title_lbl = QtWidgets.QLabel(title)
        self.title_lbl.setObjectName("title")
        self.value_lbl = QtWidgets.QLabel("-")
        self.value_lbl.setStyleSheet("font-size:28px; font-weight:800; color:#0f172a;")
        layout.addWidget(self.title_lbl)
        layout.addWidget(self.value_lbl)
        layout.addStretch(1)
        self.setLayout(layout)

    def set_value(self, value: str):
        self.value_lbl.setText(value)


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, cfg: AppConfig):
        super().__init__()
        self.cfg = cfg
        self.cfg.paths.ensure()
        db.ensure_schema(self.cfg)
        self.setWindowTitle(f"NAOS SLA Tracker v{cfg.version}")
        self.resize(1400, 900)
        self.statusBar().showMessage("Разработчик: Тёма (ShapArt)")
        self._threads: List[TaskThread] = []
        self.log_console = LogConsole()
        self.log_bridge = QtLogBridge(self.log_console.append_line)
        self.log_console.command_entered.connect(self._handle_command)
        self._build_ui()
        self._refresh_env()
        self._refresh_dashboard()
        self._refresh_table()

    # ---------- UI build ----------
    def _build_ui(self):
        splitter = QtWidgets.QSplitter(QtCore.Qt.Vertical)
        splitter.addWidget(self._build_top_panel())
        splitter.addWidget(self.log_console)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)
        self.setCentralWidget(splitter)

    def _build_top_panel(self) -> QtWidgets.QWidget:
        container = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout()
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(12)

        header = QtWidgets.QHBoxLayout()
        header.setSpacing(12)
        self.run_button = QtWidgets.QPushButton("Утро/вечер: полный цикл")
        self.run_button.setMinimumHeight(64)
        self.run_button.setStyleSheet(
            "font-size:18px; font-weight:800; padding:14px 22px; background:qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #2563eb, stop:1 #1d4ed8); color:white; border-radius:14px;"
        )
        self.run_button.clicked.connect(
            lambda: self._run_task("Полный сценарий", self._scenario_run)
        )

        header_info = QtWidgets.QVBoxLayout()
        header_info.setSpacing(6)
        self.demo_button = QtWidgets.QPushButton("Демо/QA сценарий (без отправки)")
        self.demo_button.setMinimumHeight(40)
        self.demo_button.setStyleSheet("font-weight:600;")
        self.demo_button.clicked.connect(
            lambda: self._run_task("Демо", self._demo_test)
        )
        self.link_docs = QtWidgets.QCommandLinkButton("Гайд / ссылки / открыть конфиг")
        self.link_docs.clicked.connect(self._open_docs)
        self.send_state_badge = QtWidgets.QLabel()
        self.send_state_badge.setStyleSheet(
            "padding:4px 8px; border-radius:8px; font-weight:700;"
        )
        header_info.addWidget(self.demo_button)
        header_info.addWidget(self.link_docs)
        header_info.addWidget(self.send_state_badge)

        header.addWidget(self.run_button, 2)
        header.addLayout(header_info, 1)
        layout.addLayout(header)

        metrics = QtWidgets.QHBoxLayout()
        metrics.setSpacing(12)
        self.metric_total = MetricCard("Всего заявок")
        self.metric_overdue = MetricCard("Просрочки")
        self.metric_need = MetricCard("В работе")
        metrics.addWidget(self.metric_total)
        metrics.addWidget(self.metric_overdue)
        metrics.addWidget(self.metric_need)
        layout.addLayout(metrics)

        # Filters
        self.filter_bar = QtWidgets.QHBoxLayout()
        self.filter_group = QtWidgets.QButtonGroup(self)
        filters = [
            ("all", "Все"),
            ("need", "Нужны действия"),
            ("overdue", "Просрочено"),
            ("no_owner", "Без ответственного"),
            ("conflict", "Конфликты Excel"),
        ]
        for key, label in filters:
            btn = QtWidgets.QPushButton(label)
            btn.setCheckable(True)
            if key == "need":
                btn.setChecked(True)
            self.filter_group.addButton(btn)
            self.filter_group.setId(btn, len(self.filter_group.buttons()))
            btn.clicked.connect(lambda _, k=key: self._refresh_table(k))
            self.filter_bar.addWidget(btn)
        self.filter_bar.addStretch(1)
        layout.addLayout(self.filter_bar)

        # Center responsive area
        center_widget = self._build_center_area()
        layout.addWidget(center_widget, 1)

        # Settings + info row
        info_row = QtWidgets.QHBoxLayout()
        info_row.setSpacing(12)
        info_row.addWidget(self._build_settings_panel(), 2)
        info_row.addWidget(self._build_info_panel(), 1)
        layout.addLayout(info_row)

        container.setLayout(layout)
        return container

    def _build_center_area(self) -> QtWidgets.QWidget:
        self.table = QtWidgets.QTableView()
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setDefaultSectionSize(26)
        self.table.setStyleSheet("QTableView { font-size: 11pt; }")
        self.table.clicked.connect(self._on_row_selected)
        self.model = QtGui.QStandardItemModel()
        self.table.setModel(self.model)

        self.detail_panel = self._build_detail_panel()

        self.splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        self.wide_container = QtWidgets.QWidget()
        wide_layout = QtWidgets.QVBoxLayout()
        wide_layout.setContentsMargins(0, 0, 0, 0)
        wide_layout.addWidget(self.splitter)
        self.wide_container.setLayout(wide_layout)

        self.narrow_list_container = QtWidgets.QWidget()
        self.narrow_list_layout = QtWidgets.QVBoxLayout()
        self.narrow_list_layout.setContentsMargins(0, 0, 0, 0)
        self.narrow_list_container.setLayout(self.narrow_list_layout)

        self.narrow_detail_container = QtWidgets.QWidget()
        self.narrow_detail_layout = QtWidgets.QVBoxLayout()
        self.narrow_detail_layout.setContentsMargins(0, 0, 0, 0)
        self.narrow_back_btn = QtWidgets.QPushButton("← К списку")
        self.narrow_back_btn.clicked.connect(self._back_to_list)
        self.narrow_detail_layout.addWidget(self.narrow_back_btn)
        self.narrow_detail_layout.addWidget(self.detail_panel)
        self.narrow_detail_container.setLayout(self.narrow_detail_layout)

        self.narrow_stack = QtWidgets.QStackedWidget()
        self.narrow_stack.addWidget(self.narrow_list_container)
        self.narrow_stack.addWidget(self.narrow_detail_container)

        self.center_stack = QtWidgets.QStackedWidget()
        self.center_stack.addWidget(self.wide_container)
        self.center_stack.addWidget(self.narrow_stack)
        self._current_layout_mode: Optional[str] = None
        self._set_layout_mode("wide")

        return self.center_stack

    def _detach_widget(self, widget: QtWidgets.QWidget):
        parent = widget.parent()
        if parent and isinstance(parent, QtWidgets.QWidget):
            layout = parent.layout()
            if layout:
                layout.removeWidget(widget)
        widget.setParent(None)

    def _set_layout_mode(self, mode: str):
        if mode == self._current_layout_mode:
            return
        self._detach_widget(self.table)
        self._detach_widget(self.detail_panel)
        if mode == "wide":
            self.splitter.addWidget(self.table)
            self.splitter.addWidget(self.detail_panel)
            self.splitter.setStretchFactor(0, 3)
            self.splitter.setStretchFactor(1, 2)
            self.splitter.setHandleWidth(8)
            self.center_stack.setCurrentWidget(self.wide_container)
        else:
            self.narrow_list_layout.addWidget(self.table)
            self.narrow_stack.setCurrentWidget(self.narrow_list_container)
            self.center_stack.setCurrentWidget(self.narrow_stack)
        self._current_layout_mode = mode

    def resizeEvent(self, event: QtGui.QResizeEvent):  # noqa: N802
        super().resizeEvent(event)
        self._update_layout_mode()

    def _update_layout_mode(self):
        mode = "narrow" if self.width() < NARROW_THRESHOLD else "wide"
        self._set_layout_mode(mode)

    def _build_detail_panel(self) -> QtWidgets.QWidget:
        panel = QtWidgets.QGroupBox("Карточка")
        panel.setStyleSheet("QGroupBox{font-weight:700;}")
        layout = QtWidgets.QVBoxLayout()
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(8)
        self.lbl_subject = QtWidgets.QLabel("Тема: -")
        self.lbl_sender = QtWidgets.QLabel("Отправитель: -")
        self.lbl_client = QtWidgets.QLabel("Email клиента: -")
        self.lbl_sla = QtWidgets.QLabel("SLA: -")
        self.lbl_repeat = QtWidgets.QLabel("Повторы/рекомендации: -")
        for lbl in (
            self.lbl_subject,
            self.lbl_sender,
            self.lbl_client,
            self.lbl_sla,
            self.lbl_repeat,
        ):
            lbl.setWordWrap(True)
            lbl.setStyleSheet("font-size:11pt;")
        info_grid = QtWidgets.QGridLayout()
        info_grid.setVerticalSpacing(4)
        info_grid.setHorizontalSpacing(8)
        info_grid.addWidget(self.lbl_subject, 0, 0, 1, 2)
        info_grid.addWidget(self.lbl_sender, 1, 0, 1, 2)
        info_grid.addWidget(self.lbl_client, 2, 0, 1, 2)
        info_grid.addWidget(self.lbl_sla, 3, 0, 1, 2)
        info_grid.addWidget(self.lbl_repeat, 4, 0, 1, 2)
        self.body_view = QtWidgets.QTextEdit()
        self.body_view.setReadOnly(True)
        self.body_view.setMinimumHeight(120)
        self.body_view.setStyleSheet("font-size:11pt; background:#f8fafc;")

        form = QtWidgets.QFormLayout()
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(6)
        self.status_combo = QtWidgets.QComboBox()
        for code, label_ru in sla.STATUS_MODEL:
            self.status_combo.addItem(label_ru, code)
        self.responsible_input = QtWidgets.QLineEdit()
        self.priority_combo = QtWidgets.QComboBox()
        self.priority_combo.addItems(["", "p1", "p2", "p3", "p4"])
        self.comment_input = QtWidgets.QLineEdit()
        form.addRow("Статус", self.status_combo)
        form.addRow("Ответственный", self.responsible_input)
        form.addRow("Приоритет", self.priority_combo)
        form.addRow("Комментарий", self.comment_input)

        actions = QtWidgets.QHBoxLayout()
        actions.setSpacing(8)
        btn_save = QtWidgets.QPushButton("Сохранить")
        btn_save.clicked.connect(self._save_ticket_changes)
        btn_remind = QtWidgets.QPushButton("Отправить напоминание")
        btn_remind.clicked.connect(self._send_single_reminder)
        btn_excel = QtWidgets.QPushButton("Открыть Excel")
        btn_excel.clicked.connect(self._open_excel_for_ticket)
        btn_outlook = QtWidgets.QPushButton("Открыть письмо в Outlook")
        btn_outlook.clicked.connect(self._open_outlook_item)
        for b in (btn_save, btn_remind, btn_excel, btn_outlook):
            actions.addWidget(b)
        actions.addStretch(1)

        self.events_list = QtWidgets.QListWidget()
        self.events_list.setMaximumHeight(140)
        self.events_list.setStyleSheet(
            "font-family: Consolas, 'Cascadia Code', monospace; font-size:10pt;"
        )

        layout.addLayout(info_grid)
        layout.addWidget(self.body_view)
        layout.addLayout(form)
        layout.addLayout(actions)
        layout.addWidget(QtWidgets.QLabel("История событий"))
        layout.addWidget(self.events_list)
        panel.setLayout(layout)
        return panel

    def _build_settings_panel(self) -> QtWidgets.QWidget:
        panel = QtWidgets.QGroupBox("Настройки")
        panel.setStyleSheet("QGroupBox{font-weight:700;}")
        form = QtWidgets.QGridLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(6)
        self.ingest_days_spin = QtWidgets.QSpinBox()
        self.ingest_days_spin.setRange(1, 90)
        self.ingest_days_spin.setValue(self.cfg.ingest_days or 7)
        self.safe_checkbox = QtWidgets.QCheckBox("SAFE: только предпросмотр")
        self.safe_checkbox.setChecked(self.cfg.safe_mode or not self.cfg.allow_send)
        self.safe_checkbox.setToolTip(
            "В SAFE режиме письма показываются, но не отправляются"
        )
        self.send_checkbox = QtWidgets.QCheckBox("SEND: отправлять письма")
        self.send_checkbox.setChecked(self.cfg.allow_send)
        self.send_checkbox.setToolTip(
            "Включайте только осознанно; действует allowlist доменов"
        )
        self.send_checkbox.stateChanged.connect(self._confirm_send_toggle)
        self.docs_input = QtWidgets.QLineEdit(
            self.cfg.docs_url or self.cfg.sharepoint_url or ""
        )
        self.password_input = QtWidgets.QLineEdit(self.cfg.excel_password or "naos")
        self.password_input.setEchoMode(QtWidgets.QLineEdit.Password)

        self.mailbox_input = QtWidgets.QLineEdit(self.cfg.mailbox or "")
        self.folder_input = QtWidgets.QLineEdit(self.cfg.folder or "")
        self.sender_mode_input = QtWidgets.QComboBox()
        for mode in ("off", "contains", "equals", "domain"):
            self.sender_mode_input.addItem(mode)
        self.sender_mode_input.setCurrentText(self.cfg.sender_filter_mode or "off")
        self.sender_value_input = QtWidgets.QLineEdit(
            self.cfg.sender_filter_value or self.cfg.sender_filter or ""
        )

        row = 0
        form.addWidget(QtWidgets.QLabel("Период ingest (дни)"), row, 0)
        form.addWidget(self.ingest_days_spin, row, 1)
        row += 1
        form.addWidget(self.safe_checkbox, row, 0, 1, 2)
        row += 1
        form.addWidget(self.send_checkbox, row, 0, 1, 2)
        row += 1
        form.addWidget(QtWidgets.QLabel("Гайд URL / SharePoint"), row, 0)
        form.addWidget(self.docs_input, row, 1)
        row += 1
        form.addWidget(QtWidgets.QLabel("Пароль на Excel"), row, 0)
        form.addWidget(self.password_input, row, 1)
        row += 1
        form.addWidget(QtWidgets.QLabel("Почтовый ящик (mailbox)"), row, 0)
        form.addWidget(self.mailbox_input, row, 1)
        row += 1
        form.addWidget(QtWidgets.QLabel("Папка Outlook (Inbox/Support)"), row, 0)
        form.addWidget(self.folder_input, row, 1)
        row += 1
        form.addWidget(
            QtWidgets.QLabel("Фильтр отправителя (off/contains/equals/domain)"), row, 0
        )
        form.addWidget(self.sender_mode_input, row, 1)
        row += 1
        form.addWidget(QtWidgets.QLabel("Значение фильтра"), row, 0)
        form.addWidget(self.sender_value_input, row, 1)

        panel.setLayout(form)
        return panel

    def _build_info_panel(self) -> QtWidgets.QWidget:
        panel = QtWidgets.QGroupBox("Инфо")
        layout = QtWidgets.QVBoxLayout()
        layout.setSpacing(8)
        self.lbl_paths = QtWidgets.QLabel(
            f"Excel: {self.cfg.paths.excel_path}\nДанные: {self.cfg.paths.appdata_dir}\nConfig: {self.cfg.paths.appdata_dir/'config.json'}"
        )
        self.lbl_paths.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        btn_open_excel = QtWidgets.QPushButton("Открыть Excel")
        btn_open_excel.clicked.connect(self._open_excel_path)
        btn_open_data = QtWidgets.QPushButton("Открыть папку данных/логи")
        btn_open_data.clicked.connect(self._open_data_dir)
        btn_open_config = QtWidgets.QPushButton("Открыть config.json")
        btn_open_config.clicked.connect(self._open_config)
        btn_open_runbook = QtWidgets.QPushButton("Показать гайд (встроенный)")
        btn_open_runbook.clicked.connect(self._show_help_dialog)
        self.env_label = QtWidgets.QLabel("")
        self.env_label.setWordWrap(True)
        self.env_label.setStyleSheet("color:#0f172a; font-weight:600;")

        layout.addWidget(self.lbl_paths)
        layout.addWidget(btn_open_excel)
        layout.addWidget(btn_open_data)
        layout.addWidget(btn_open_config)
        layout.addWidget(btn_open_runbook)
        layout.addWidget(self.env_label)
        layout.addStretch(1)
        panel.setLayout(layout)
        return panel

    # ---------- Actions ----------
    def _handle_command(self, cmd: str):
        cmd_l = cmd.lower()
        if cmd_l in ("help", "?"):
            self.log_console.append_line(
                "Команды: sync, ingest, export, send, process, diagnose, open excel"
            )
        elif cmd_l == "sync":
            self._run_task("Sync Excel", self._sync_excel)
        elif cmd_l == "ingest":
            self._run_task("Ingest", self._ingest)
        elif cmd_l == "export":
            self._run_task("Export Excel", self._export_excel)
        elif cmd_l == "send":
            self._run_task("Send overdue", self._send_overdue)
        elif cmd_l == "process":
            self._run_task("Process responses", self._process_responses)
        elif cmd_l == "diagnose":
            self._run_task("Diagnose Outlook", self._diagnose)
        elif cmd_l == "open excel":
            self._run_task("Open Excel", self._open_excel_path)
        else:
            self.log_console.append_line("Неизвестная команда")

    def _run_task(self, label: str, fn: Callable[[], object]):
        thread = TaskThread(label, fn)
        thread.finished_ok.connect(self._on_task_ok)
        thread.finished_err.connect(self._on_task_err)
        thread.result_ready.connect(lambda _: self._refresh_dashboard())
        thread.finished.connect(
            lambda: self._threads.remove(thread) if thread in self._threads else None
        )
        self._threads.append(thread)
        thread.start()

    def _scenario_run(self):
        cfg = self._cfg_from_ui()
        days = cfg.ingest_days or 7
        self._emit_ui_log("Шаг 1: Sync Excel -> DB")
        sync_res = excel.sync_from_excel(cfg)
        self._emit_ui_log(f"Sync Excel: {sync_res}")
        self._emit_ui_log(f"Шаг 2: Ingest писем за {days} дней")
        ingested = sla.ingest_range(cfg, days=days)
        self._emit_ui_log(f"Ingest: {ingested}")
        self._emit_ui_log("Шаг 3: Пересчет SLA")
        updated = sla.recalc_open(cfg)
        self._emit_ui_log(f"Пересчитано: {updated}")
        self._emit_ui_log("Шаг 4: Export Excel")
        export_path = excel.export_excel(
            cfg,
            conflicts=(
                sync_res.get("conflict_rows") if isinstance(sync_res, dict) else None
            ),
        )
        self._emit_ui_log(f"Excel выгружен: {export_path}")
        self._emit_ui_log("Шаг 5: Process responses")
        processed = sla.process_responses(cfg, days=days)
        self._emit_ui_log(f"Ответов обработано: {processed}")
        self._emit_ui_log("Шаг 6: Send overdue (preview если SAFE)")
        plan = sla.overdue_plan(cfg)
        sent = 0
        for row in plan["send"]:
            notify.send_overdue_mail(
                cfg,
                row["responsible"],
                row["subject"],
                row["id"],
                priority=row.get("priority"),
                original_entry_id=row.get("entry_id"),
            )
            sla.mark_reminder_sent(cfg, row["id"])
            sent += 1
        summary = {
            "sync": sync_res,
            "ingested": ingested,
            "recalc": updated,
            "processed": processed,
            "sent": sent,
            "skip_interval": len(plan["skip_interval"]),
            "skip_quiet": len(plan["skip_quiet"]),
            "skip_no_owner": len(plan["skip_no_responsible"]),
        }
        self._emit_ui_log(f"Итог: {summary}")
        return summary

    def _demo_test(self):
        cfg = self._cfg_from_ui()
        conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
        ticket_id = db.seed_test_ticket(conn, status=sla.STATUS_OVERDUE)
        conn.close()
        notify.send_test_mail(
            cfg,
            "[TEST][SLA] Просрочка",
            "Сценарий демо, нажмите Voting 'Закрыть' или /status closed.",
        )
        excel.export_excel(cfg, today_only=False)
        return f"Создан тестовый тикет #{ticket_id} и отправлен тестовый email"

    def _sync_excel(self):
        cfg = self._cfg_from_ui()
        return excel.sync_from_excel(cfg)

    def _ingest(self):
        cfg = self._cfg_from_ui()
        return sla.ingest_range(cfg, days=cfg.ingest_days or 7)

    def _export_excel(self):
        cfg = self._cfg_from_ui()
        return excel.export_excel(cfg, today_only=False)

    def _send_overdue(self):
        cfg = self._cfg_from_ui()
        plan = sla.overdue_plan(cfg)
        sent = 0
        for row in plan["send"]:
            notify.send_overdue_mail(
                cfg,
                row["responsible"],
                row["subject"],
                row["id"],
                priority=row.get("priority"),
                original_entry_id=row.get("entry_id"),
            )
            sla.mark_reminder_sent(cfg, row["id"])
            sent += 1
        return {
            "sent": sent,
            "skip_interval": len(plan["skip_interval"]),
            "skip_quiet": len(plan["skip_quiet"]),
            "skip_no_responsible": len(plan["skip_no_responsible"]),
        }

    def _process_responses(self):
        cfg = self._cfg_from_ui()
        return sla.process_responses(cfg, days=cfg.ingest_days or 7)

    def _diagnose(self):
        cfg = self._cfg_from_ui()
        env = detect_outlook_environment()
        env_text = f"Classic Outlook: {'OK' if env.classic_available else 'НЕТ'}; New: {env.new_outlook_detected}; {env.details}"
        with OutlookClient(cfg) as outlook:
            before, after, top10 = outlook.diagnose(cfg.ingest_days or 7)
        top_str = (
            ", ".join(f"{s or '<unknown>'}:{c}" for s, c in top10) if top10 else "-"
        )
        return f"{env_text}; до фильтра={before}, после фильтра={after}; топ отправители: {top_str}"

    def _open_excel_path(self):
        path = self.cfg.paths.excel_path
        path.parent.mkdir(parents=True, exist_ok=True)
        if not path.exists():
            excel.export_excel(self.cfg, today_only=False)
        os.startfile(path)
        return str(path)

    def _open_data_dir(self):
        path = self.cfg.paths.appdata_dir
        path.mkdir(parents=True, exist_ok=True)
        os.startfile(path)
        return str(path)

    def _open_config(self):
        cfg_path = self.cfg.paths.appdata_dir / "config.json"
        cfg_path.parent.mkdir(parents=True, exist_ok=True)
        if not cfg_path.exists():
            self.cfg.save(cfg_path)
        os.startfile(cfg_path)
        return str(cfg_path)

    def _open_docs(self):
        url = (
            self.docs_input.text().strip()
            or self.cfg.sharepoint_url
            or self.cfg.docs_url
        )
        if url:
            QDesktopServices.openUrl(QUrl(url))
            return url
        return self._show_help_dialog()

    def _show_help_dialog(self):
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("Гайд")
        dialog.resize(800, 600)
        layout = QtWidgets.QVBoxLayout(dialog)
        browser = QtWidgets.QTextBrowser()
        browser.setOpenExternalLinks(True)
        content = None
        base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
        for name in (
            "RUNBOOK.md",
            "README.md",
            os.path.join("qa", "runbook_e2e_artem.md"),
        ):
            candidate = base / name
            if candidate.exists():
                content = candidate.read_text(encoding="utf-8")
                break
        if content:
            browser.setMarkdown(content)
        else:
            browser.setPlainText("README/RUNBOOK не найдены")
        btn_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        btn_box.rejected.connect(dialog.reject)
        layout.addWidget(browser)
        layout.addWidget(btn_box)
        dialog.exec()
        return True

    def _on_task_ok(self, msg: str):
        self.log_console.append_line(msg)
        self._refresh_dashboard()
        self._refresh_table()

    def _on_task_err(self, msg: str):
        self.log_console.append_line(msg)
        self._refresh_dashboard()
        self._refresh_table()

    def _emit_ui_log(self, text: str):
        self.log_console.append_line(f"[UI] {text}")
        log.info(text)

    # ---------- Dashboard ----------
    def _refresh_env(self):
        env = detect_outlook_environment()
        send_mode = (
            "SAFE (Display)"
            if (self.cfg.safe_mode or not self.cfg.allow_send)
            else "SEND"
        )
        text = f"Classic Outlook COM: {'OK' if env.classic_available else 'НЕТ'}; New Outlook: {env.new_outlook_detected}; {env.details}; Режим: {send_mode}"
        self.log_console.append_line(text)
        badge_style = (
            "color:#fff; background:#16a34a;"
            if send_mode.startswith("SAFE")
            else "color:#fff; background:#dc2626;"
        )
        badge_text = (
            "SEND OFF (предпросмотр)"
            if send_mode.startswith("SAFE")
            else "SEND ON (боевой режим)"
        )
        self.send_state_badge.setText(badge_text)
        self.send_state_badge.setStyleSheet(
            self.send_state_badge.styleSheet() + badge_style
        )
        self.env_label.setText(text)

    def _refresh_dashboard(self):
        try:
            conn = db.connect(self.cfg.paths.db_path, wal_mode=self.cfg.wal_mode)
            total = conn.execute("SELECT count(*) FROM tickets").fetchone()[0]
            overdue = conn.execute(
                "SELECT count(*) FROM tickets WHERE overdue=1"
            ).fetchone()[0]
            need = conn.execute(
                "SELECT count(*) FROM tickets WHERE status IN ('new','assigned','table','otip','waiting_customer')"
            ).fetchone()[0]
            conn.close()
            self.metric_total.set_value(str(total))
            self.metric_overdue.set_value(str(overdue))
            self.metric_need.set_value(str(need))
        except Exception as exc:
            log.warning("dashboard refresh failed: %s", exc)

    # ---------- Table ----------
    def _refresh_table(self, filter_key: str = "need"):
        headers = [
            "ID",
            "Статус",
            "Ответственный",
            "Приоритет",
            "Тема",
            "Клиент",
            "Получено",
            "Дней без обновления",
            "Повторы",
            "Конфликт",
        ]
        self.model.clear()
        self.model.setHorizontalHeaderLabels(headers)
        try:
            conn = db.connect(self.cfg.paths.db_path, wal_mode=self.cfg.wal_mode)
            where = []
            params: List = []
            if filter_key == "need":
                where.append(
                    "status IN ('new','assigned','table','otip','waiting_customer')"
                )
            elif filter_key == "overdue":
                where.append("overdue=1 OR status='overdue'")
            elif filter_key == "no_owner":
                where.append("(responsible IS NULL OR responsible='')")
            base_sql = """
            SELECT t.*, EXISTS(SELECT 1 FROM events e WHERE e.ticket_id=t.id AND e.event_type='excel_conflict') AS has_conflict
            FROM tickets t
            """
            if filter_key == "conflict":
                base_sql += " WHERE EXISTS(SELECT 1 FROM events e WHERE e.ticket_id=t.id AND e.event_type='excel_conflict')"
            elif where:
                base_sql += " WHERE " + " AND ".join(where)
            base_sql += " ORDER BY datetime(first_received_utc) DESC LIMIT 500"
            rows = conn.execute(base_sql, params).fetchall()
            self._current_rows: Dict[int, dict] = {}
            for r in rows:
                rid = int(r["id"])
                status_label = sla.status_code_to_label(r["status"])
                responsible = r["responsible"] or ""
                priority = r["priority"] or ""
                subject = r["subject"]
                customer = r["customer_email"] or ""
                sla_due = r["first_received_utc"]
                days_wo = r["days_without_update"]
                repeat = r["repeat_hint"] or ""
                conflict = "??" if r["has_conflict"] else ""
                items = [
                    QtGui.QStandardItem(str(rid)),
                    QtGui.QStandardItem(status_label),
                    QtGui.QStandardItem(responsible),
                    QtGui.QStandardItem(priority),
                    QtGui.QStandardItem(subject),
                    QtGui.QStandardItem(customer),
                    QtGui.QStandardItem(sla_due),
                    QtGui.QStandardItem(str(days_wo)),
                    QtGui.QStandardItem(repeat),
                    QtGui.QStandardItem(conflict),
                ]
                for item in items:
                    item.setEditable(False)
                if r["overdue"]:
                    for item in items:
                        item.setBackground(QtGui.QColor("#FEE2E2"))
                elif r["status"] in (
                    sla.STATUS_NEW,
                    sla.STATUS_ASSIGNED,
                    sla.STATUS_TABLE,
                ):
                    for item in items:
                        item.setBackground(QtGui.QColor("#EFF6FF"))
                self.model.appendRow(items)
                self._current_rows[rid] = dict(r)
            conn.close()
        except Exception as exc:
            log.warning("table refresh failed: %s", exc)

    def _on_row_selected(self, index: QtCore.QModelIndex):
        if not index.isValid():
            return
        rid_item = self.model.item(index.row(), 0)
        if not rid_item:
            return
        try:
            ticket_id = int(rid_item.text())
        except Exception:
            return
        self._load_ticket_details(ticket_id)
        if self._current_layout_mode == "narrow":
            self.narrow_stack.setCurrentWidget(self.narrow_detail_container)

    def _back_to_list(self):
        if self._current_layout_mode == "narrow":
            self.narrow_stack.setCurrentWidget(self.narrow_list_container)

    def _load_ticket_details(self, ticket_id: int):
        try:
            conn = db.connect(self.cfg.paths.db_path, wal_mode=self.cfg.wal_mode)
            row = conn.execute(
                "SELECT * FROM tickets WHERE id=?", (ticket_id,)
            ).fetchone()
            events = conn.execute(
                "SELECT event_type, status_before, status_after, source, event_dt_utc FROM events WHERE ticket_id=? ORDER BY datetime(event_dt_utc) DESC LIMIT 30",
                (ticket_id,),
            ).fetchall()
            conn.close()
            if not row:
                return
            self.current_ticket_id = ticket_id
            self.lbl_subject.setText(f"Тема: {row['subject']}")
            self.lbl_sender.setText(f"Отправитель: {row['sender']}")
            self.lbl_client.setText(f"Email клиента: {row['customer_email'] or '-'}")
            self.lbl_sla.setText(
                f"Статус: {sla.status_code_to_label(row['status'])} | Просрочка: {'Да' if row['overdue'] else 'Нет'} | Дней без обновления: {row['days_without_update']}"
            )
            self.lbl_repeat.setText(
                f"Повторы/советы: {row['repeat_hint'] or row['recommended_answer'] or '-'}"
            )
            self.body_view.setPlainText(row["body"])
            idx = self.status_combo.findData(row["status"])
            self.status_combo.setCurrentIndex(idx if idx >= 0 else 0)
            self.responsible_input.setText(row["responsible"] or "")
            prio_idx = self.priority_combo.findText(row["priority"] or "")
            self.priority_combo.setCurrentIndex(prio_idx if prio_idx >= 0 else 0)
            self.comment_input.setText(row["comment"] or "")
            self.events_list.clear()
            for ev in events:
                self.events_list.addItem(
                    f"{ev['event_dt_utc']} [{ev['source']}::{ev['event_type']}] {ev['status_before']}->{ev['status_after']}"
                )
        except Exception as exc:
            log.warning("load details failed: %s", exc)

    def _save_ticket_changes(self):
        if not getattr(self, "current_ticket_id", None):
            return
        cfg = self._cfg_from_ui()
        ticket_id = self.current_ticket_id
        new_status = self.status_combo.currentData()
        resp = self.responsible_input.text().strip()
        prio = self.priority_combo.currentText().strip()
        comment = self.comment_input.text().strip()
        try:
            conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
            conn.execute(
                """
                UPDATE tickets
                SET status=?, responsible=?, priority=?, comment=?, days_without_update=0,
                    overdue=CASE WHEN ? IN ('resolved','table','not_interesting') THEN 0 ELSE overdue END,
                    last_status_utc=datetime('now'), last_updated_at=datetime('now'), last_updated_by='ui', row_version=row_version+1, data_source='ui'
                WHERE id=?
                """,
                (
                    new_status,
                    resp or None,
                    prio or None,
                    comment or None,
                    new_status,
                    ticket_id,
                ),
            )
            db.log_event(conn, ticket_id, "ui_update", None, new_status, "ui")
            conn.commit()
            conn.close()
            self.log_console.append_line(
                f"Saved #{ticket_id}: {new_status}, {resp}, {prio}"
            )
            self._refresh_table()
        except Exception as exc:
            log.warning("save failed: %s", exc)

    def _send_single_reminder(self):
        if not getattr(self, "current_ticket_id", None):
            return
        cfg = self._cfg_from_ui()
        plan = sla.overdue_plan(cfg)
        found = None
        for bucket, rows in plan.items():
            for r in rows:
                if r.get("id") == self.current_ticket_id:
                    found = (bucket, r)
                    break
        if not found:
            self.log_console.append_line(
                "No reminder candidate (quiet/no responsible/interval)"
            )
            return
        bucket, row = found
        if bucket != "send":
            reason = {
                "skip_interval": "interval",
                "skip_quiet": "quiet hours",
                "skip_no_responsible": "no responsible",
            }.get(bucket, bucket)
            self.log_console.append_line(f"Skipped: {reason}")
            return
        notify.send_overdue_mail(
            cfg,
            row["responsible"],
            row["subject"],
            row["id"],
            priority=row.get("priority"),
            original_entry_id=row.get("entry_id"),
        )
        sla.mark_reminder_sent(cfg, row["id"])
        self.log_console.append_line(f"Reminder sent for #{row['id']}")

    def _open_excel_for_ticket(self):
        self._open_excel_path()

    def _open_outlook_item(self):
        if not getattr(self, "current_ticket_id", None):
            return
        cfg = self._cfg_from_ui()
        conn = db.connect(cfg.paths.db_path, wal_mode=cfg.wal_mode)
        row = conn.execute(
            "SELECT entry_id FROM tickets WHERE id=?", (self.current_ticket_id,)
        ).fetchone()
        conn.close()
        entry_id = row["entry_id"] if row else None
        if not entry_id:
            self.log_console.append_line("No entry_id to open in Outlook")
            return
        try:
            with OutlookClient(cfg) as outlook:
                item = outlook.ns.GetItemFromID(entry_id)
                item.Display()
        except Exception as exc:
            self.log_console.append_line(f"Failed to open item: {exc}")

    # ---------- Helpers ----------
    def _confirm_send_toggle(self, state: int):
        if state == QtCore.Qt.Checked:
            confirm = QtWidgets.QMessageBox.question(
                self,
                "Enable SEND?",
                "Sending will be enabled. Make sure allowlist/domains are configured. Enable?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
            )
            if confirm != QtWidgets.QMessageBox.Yes:
                self.send_checkbox.blockSignals(True)
                self.send_checkbox.setChecked(False)
                self.send_checkbox.blockSignals(False)
        self._update_send_state_ui()

    def _update_send_state_ui(self):
        badge_style = (
            "color:#fff; background:#16a34a;"
            if (not self.send_checkbox.isChecked() or self.safe_checkbox.isChecked())
            else "color:#fff; background:#dc2626;"
        )
        badge_text = (
            "SEND OFF (preview)"
            if (not self.send_checkbox.isChecked() or self.safe_checkbox.isChecked())
            else "SEND ON (live)"
        )
        self.send_state_badge.setText(badge_text)
        self.send_state_badge.setStyleSheet(
            self.send_state_badge.styleSheet() + badge_style
        )

    def _cfg_from_ui(self) -> AppConfig:
        cfg = self.cfg
        cfg.ingest_days = int(self.ingest_days_spin.value())
        cfg.safe_mode = self.safe_checkbox.isChecked()
        cfg.allow_send = self.send_checkbox.isChecked()
        cfg.docs_url = self.docs_input.text().strip() or None
        cfg.sharepoint_url = cfg.docs_url
        cfg.excel_password = self.password_input.text().strip() or "naos"
        cfg.mailbox = self.mailbox_input.text().strip() or None
        cfg.folder = self.folder_input.text().strip() or None
        cfg.sender_filter_mode = self.sender_mode_input.currentText()
        cfg.sender_filter_value = self.sender_value_input.text().strip()
        cfg.sender_filter = cfg.sender_filter_value
        cfg.save()
        self._update_send_state_ui()
        return cfg


def run_ui(cfg: AppConfig | None = None):
    app = QtWidgets.QApplication(sys.argv)
    if cfg is None:
        cfg = default_config()
    cfg.paths.ensure()
    setup_logging(cfg.paths.log_dir)
    win = MainWindow(cfg)
    win.show()
    sys.exit(app.exec())
