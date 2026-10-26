# lab-report-utm

Скилл для Claude Code: пишет и оформляет отчёты по лабораторным работам для UTM (FCIM) — готовый `.docx` + PDF с титульным листом, оглавлением, нумерацией, рисунками, таблицами, листингами и библиографией по нормам технического редактирования.

## Установка

Положи папку в каталог скиллов Claude Code:

```bash
git clone https://github.com/<user>/lab-report-utm "$HOME/.claude/skills/lab-report-utm"
```

(Windows PowerShell: `git clone https://github.com/<user>/lab-report-utm "$env:USERPROFILE\.claude\skills\lab-report-utm"`)

## Настройка

Открой `assets/profile.json` и впиши свои ФИО (ru / ro / en) и группу. Остальное (дисциплина, тема, номер работы, преподаватель) берётся из задания.

## Использование

В Claude Code напиши `/lab-report-utm` и приложи задание, код и скриншоты — или просто попроси «оформи лабу».

## Требования

- Python 3 с `python-docx` (см. `scripts/build_report.py`)
- Microsoft Word на Windows — для рендера PDF (`scripts/render_docx.ps1`)
