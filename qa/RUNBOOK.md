# QA RUNBOOK (коротко)

1. `python cli.py qa-full --send` (или без --send для Display):
   - гоняет pytest
   - запускает qa/tools/qa_e2e_driver.py (safe по умолчанию)
   - пишет отчет `%APPDATA%/NAOS_SLA_TRACKER/QA_REPORT.md`
2. Если qa_e2e_driver ждёт ручного шага — откройте письмо в Outlook, нажмите Voting "Закрыть", вернитесь и нажмите Enter.
3. Проверить QA_REPORT.md (exit_code=0 и статус тикета resolved/overdue=0, Excel путь указан).
4. В UI для визуальной проверки: запустить `python cli.py ui`, нажать "Запустить сценарий", убедиться, что чат-лог отражает шаги.
