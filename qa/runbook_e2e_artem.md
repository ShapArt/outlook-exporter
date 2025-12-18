# Golden demo: Артём → Просрочка → Ремайндер → Voting → Статус обновился

## Подготовка

- Classic Outlook запущен.
- Конфиг: `qa_artem_email` (или test_allowlist[0]) заполнен, `qa_force_overdue_days` >= 3, `qa_subject_prefix="[QA][SLA]"`.
- safe_mode=true для демо (Display писем), allow_send=true для боевого.

## Шаги

1. `python qa/tools/qa_e2e_driver.py --send` (или без --send для Display).
   - создает просроченный тикет, назначает Артёма, отправляет ремайндер reply-based (если entry_id есть).
   - ждёт ручного шага.
2. Открыть письмо в Outlook, нажать Voting "Закрыть" (или ответить `/status closed`).
3. Вернуться в консоль, нажать Enter (qa_e2e_driver продолжит process_responses, recalc, export).
4. Проверить вывод: `status=resolved`, `overdue=0`, путь к Excel.
5. Открыть Excel (tickets.xlsx) — строка тикета должна иметь Статус "Закрыто".
6. При повторном запуске send-overdue для того же тикета — писем не уйдёт (interval/quiet hours), в логе видно skip_reason.

## Диагностика

- Если COM/Outlook не найден: переключиться на Classic Outlook (new Outlook не поддерживает COM automation).
- Если ответ не найден: убедиться, что напоминание ушло как Reply (original_entry_id) и VotingResponse сохранилась.
