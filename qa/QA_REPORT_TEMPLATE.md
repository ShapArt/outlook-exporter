# QA Report Template

- Дата/время: {{timestamp}}
- Python/Windows: {{python_version}} / {{platform}}
- Outlook: classic={{classic}}, new={{new_outlook}}, details={{details}}
- Config: safe_mode={{safe_mode}}, allow_send={{allow_send}}, mailbox={{mailbox}}, folder={{folder}}

## Pytest

exit_code={{pytest_exit}}

```
{{pytest_output}}
```

## Semi E2E (qa_e2e_driver)

exit_code={{e2e_exit}}

```
{{e2e_output}}
```

## Известные риски

- New Outlook не поддерживает COM automation, требуется Classic Outlook (см. https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/).
- Object Model Guard/антивирус может блокировать win32com.Send (нужен интерактивный доступ).
- Если Excel был открыт пользователем во время экспорта — создаётся \*\_pending.xlsx.
- Если WAL недоступен — будет предупреждение, но работа продолжается.
