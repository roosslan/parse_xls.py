# parse_xls.py

Скрипт еженедельной статистики по DEV-сервисам на основе книги Excel.

## Что делает скрипт

1. Обновляет внешние запросы в `C:\Documents\Statistik Loader.xlsx` через Excel (COM).
2. Читает лист `Quelldaten`, удаляет закрытые заявки и заявки не из очередей
   `[1] Second Level::DEV` и `[2] Third Level::DEV`.
3. Считает заявки по сервисам `[T000] DEV::*` (столбец I).
4. Сохраняет листы `Quelldaten` и `T000` в `\\0x2srv\DFS-Abteilungen$\Developers\b-end\ДД.ММ.ГГГГ.xlsx`.
5. Открывает в Outlook подготовленное письмо для проверки и отправки.

Пути, получатель письма и список сервисов задаются константами в начале `parse_xls.py`.

## Развёртывание

Workflow `.github/workflows/deploy.yml`:

- `test` — запускает тесты на каждый push и pull request в `trunk`;
- `deploy` — после push в `trunk` (или ручного запуска) копирует `parse_xls.py`
  и `requirements.txt` в директорию `DEPLOY_DIR` на рабочей машине, создаёт там
  `.venv` и устанавливает зависимости. Деплой выполняется в environment `production`
  и отображается в разделе Deployments репозитория.

Для работы деплоя необходимо:

1. Зарегистрировать на рабочей машине self-hosted runner с метками `self-hosted, windows`
   (Settings → Actions → Runners). На машине должен быть Python в `PATH`.
2. В Settings → Environments → `production` задать переменную `DEPLOY_DIR`
   (например, `C:\Tools\parse_xls`) и включить Required reviewers, чтобы каждый деплой
   подтверждался вручную. Репозиторий публичный, поэтому деплой из pull request
   намеренно отключён.

После деплоя:

```bash
C:\Tools\parse_xls\.venv\Scripts\python.exe C:\Tools\parse_xls\parse_xls.py
```
