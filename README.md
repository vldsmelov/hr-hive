# HR Hive

Переносимая проверка резюме: локальные парсинг и OCR в Docker, семантическая оценка через Codex в VS Code, SQLite и Excel. Поддерживаются Mac с Apple Silicon и Intel: Compose выбирает родную архитектуру автоматически.

## Первый запуск на Mac

Установите Git, Docker Desktop для своего процессора, VS Code и расширение Codex. Запустите Docker Desktop и войдите в Codex в VS Code.

Для клонирования приватного репозитория войдите в GitHub через VS Code либо GitHub CLI (`gh auth login`). Если команда `code` не установлена, откройте папку проекта через меню VS Code → File → Open Folder.

```sh
git clone https://github.com/vldsmelov/hr-hive.git
cd hr-hive
mkdir -p data/backups
```

Перенесите **архив данных и файл `.sha256`** в `data/backups/`. CV пока не нужны. Подставьте фактическое имя архива:

```sh
cd data/backups
shasum -a 256 -c hr-hive-transfer.tar.gz.sha256
cd ../..
docker compose build
docker compose run --rm app restore --file /data/backups/hr-hive-transfer.tar.gz
docker compose run --rm app doctor
docker compose run --rm app status --all
docker compose run --rm app export
code .
```

`restore` принимает только пустое хранилище; существующие `cvs` и `backups` допустимы. Проверка всех файлов архива выполняется до записи. Не распаковывайте архив поверх работающей базы вручную.

В чате Codex в VS Code напишите: «Прочитай AGENTS.md, проверь status и продолжи обработку незавершённых резюме». Все правила сохранены в проекте и источниках версии, прежняя переписка не требуется.

## Рабочие каталоги

| Каталог | Содержимое |
|---|---|
| `data/cvs/` | Добавляемые пользователем CV, допускаются вложенные папки; имя без расширения — ID |
| `data/sources/` | Исходная книга, образец, `clusters/*.xlsx`, `banlist/*.csv` |
| `data/state/runs/` | Рабочие SQLite-базы, версии источников, карточки и журналы |
| `data/state/history/` | Архив прежнего проекта и всех запусков; старые скрипты не запускать |
| `data/exports/` | Excel и предыдущие проверенные версии |
| `data/backups/` | Переносимые архивы и контрольные суммы |

Исходные данные находятся только в локальном архиве, в GitHub их нет. SQLite-архивы не содержат бинарных копий CV; полные извлечённые тексты и история сохранены. Отсутствие исходных CV не отменяет уже завершённые проверки. После добавления CV повторно используется хеш содержимого, поэтому перенос в другую папку или Unicode-форма имени на Mac не меняют результат.

## Обычная работа

```sh
docker compose run --rm app scan
docker compose run --rm app extract
docker compose run --rm app retry --ids 123 456
docker compose run --rm app claim --worker codex --count 5
docker compose run --rm app show --id 123 --offset 0 --limit 10000
docker compose run --rm app profiles
docker compose run --rm app preview --file /data/state/cards/123.jsonl
docker compose run --rm app submit --file /data/state/cards/123.jsonl
docker compose run --rm app export
```

Для выбора другого сохранённого запуска: `docker compose run --rm app use --run ИМЯ_ИЗ_STATUS`. Активный запуск после переноса — последняя отдельная папка CV; основная база также доступна через `status --all`.

Параметры `--worker` в claim/preview/submit должны совпадать. Показ текста возвращает `next_offset`; читайте до null. Оценка не выполняется автоматически по ключевым словам. Codex читает документы полностью и составляет карточку; отдельно оплачиваемый модельный API не используется.

Каждый submit сохраняет завершённую проверку транзакцией. Выгрузка выполняется каждые 25 кандидатов или при первом сохранении после 10 минут с предыдущей выгрузки. При завершении работы Codex вызывает `export`; после аварийной остановки эта же команда полностью восстанавливает книгу из базы. Между командами нет фонового процесса или таймера.

Книга создаётся из исходного шаблона точечными изменениями XML. Все строки B–AC и их гиперссылки проверяются по исходнику; меняются A, AD–AG. Если файл занят или недоступен, создаётся отдельная версия. Ошибки распознавания сохраняются отдельно от завершённых кандидатов в базе и выводе retry/status.

Изменив правила или бан-листы в `data/sources`, выполните:

```sh
docker compose run --rm app scan --new-version
```

Старая версия сохранится; новая повторно использует тексты, а решения пересматриваются заново. Новая дата проверки также фиксируется при создании версии, что важно для срочных банов. Не исправляйте снимки критериев внутри существующего запуска.

## Резервная копия и обновления

```sh
docker compose run --rm app export
docker compose run --rm app backup --output /data/backups/hr-hive-backup.tar.gz
git pull --ff-only
docker compose build
docker compose run --rm app doctor
```

Архив содержит SQLite-снимки и контрольные суммы каждого файла. CV и изображения страниц в архив не включаются. Для возвращения к прежней копии используйте отдельную пустую папку проекта с тем же кодом.

## Проверки

```sh
docker compose run --rm -v "$PWD/tests:/tests:ro" --entrypoint python app -m unittest discover -s /tests -v
```

Набор создаёт только синтетические документы. Отчёт проверок перед передачей находится в `docs/validation.md`; локальные отчёты миграции — в `data/state/`.

Приложение поддерживает PDF, DOC/DOCX, DOTX, RTF, изображения, TXT/MD, ODT и PPTX. Pages конвертируется LibreOffice, если конкретная версия формата поддержана; иначе сохраняется явная ошибка. Аудио и незавершённые загрузки автоматически не отмечаются прочитанными. Зашифрованные/повреждённые документы требуют исправленного оригинала.
