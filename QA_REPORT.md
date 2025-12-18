# QA Report 2025-12-14T13:56:37.379373

- Python: 3.11.9 (tags/v3.11.9:de54cf5, Apr 2 2024, 10:12:12) [MSC v.1938 64 bit (AMD64)]
- Outlook: classic_available=True, new_outlook=False, details=COM ok, version=16.0.0.16327
- Config: safe_mode=True, allow_send=False, mailbox=None, folder=None

## Pytest

exit_code=0

```
============================= test session starts =============================
platform win32 -- Python 3.11.9, pytest-9.0.2, pluggy-1.6.0
rootdir: C:\Users\darks\OneDrive\Документы\Outlook
plugins: anyio-4.8.0
collected 40 items

tests\test_business_hours.py ....                                        [ 10%]
tests\test_customer_email.py .....                                       [ 22%]
tests\test_db_constraints.py ..                                          [ 27%]
tests\test_excel_features.py ...                                         [ 35%]
tests\test_excel_protection.py .                                         [ 37%]
tests\test_excel_roundtrip.py .                                          [ 40%]
tests\test_reminder_interval.py .                                        [ 42%]
tests\test_sender_filter.py ....                                         [ 52%]
tests\test_sender_filter_modes.py ....                                   [ 62%]
tests\test_sla.py ......                                                 [ 77%]
tests\test_sla_outlook_fake.py ..                                        [ 82%]
tests\test_status_mapping.py ....                                        [ 92%]
tests\test_status_text_extended.py .                                     [ 95%]
tests\test_utils.py ..                                                   [100%]

============================= 40 passed in 10.33s =============================

```

## Semi E2E (qa_e2e_driver)

exit_code=0

```
Outlook env: classic=True, new=False, details=COM ok, version=16.0.0.16327
Seeded overdue ticket #293, responsible=artem.shapovalov@ru.naos.com
Напоминание отправлено (preview/display если safe_mode).
Process responses updated: 0
Ticket after responses: status=overdue, overdue=1, row_version=1
Excel exported: C:\Users\darks\AppData\Roaming\NAOS_SLA_TRACKER\tickets.xlsx

```
