# adbeam_wb_marketguru_parser

Лёгкий отдельный проект для research MVP по задаче:

MarketGuru / Excel -> Wildberries -> реквизиты продавца -> ИНН -> новый Excel.

## Что уже есть на первом шаге

- анализ структуры входного Excel;
- подготовка нормализованной research-выборки;
- построение candidate WB URL по `Артикул`;
- Playwright smoke/research режим для открытия карточки и сохранения артефактов;
- базовый каркас CLI без лишней архитектуры.

## Что НЕ обещает текущий шаг

Текущий шаг не гарантирует стабильное извлечение ИНН.
Сейчас задача шага — быстро проверить входные данные и получить артефакты для live research:

- какие строки брать;
- открывается ли карточка по `Артикул`;
- что реально лежит в DOM;
- как потом точнее строить extractor.

## Команды

### 1. Установка зависимостей

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m pip install -e .
python -m playwright install chromium
```

### 2. Анализ Excel

```powershell
python -m wb_inn_extractor.cli analyze --input "C:\path\input.xlsx"
```

### 3. Подготовка research-выборки

```powershell
python -m wb_inn_extractor.cli sample --input "C:\path\input.xlsx" --limit 30
```

По умолчанию файл будет создан в `output\research_sample.xlsx`.

### 4. Smoke/research по одной строке

```powershell
python -m wb_inn_extractor.cli inspect-row --input "C:\path\output\research_sample.xlsx" --row 2 --headful
```

Будут сохранены:

- screenshot;
- html snapshot;
- json с первичными результатами regex-поиска реквизитов.

## Текущая логика выбора поля для WB

Во входной выгрузке нет прямого URL Wildberries.
На первом шаге основной кандидат для перехода на WB — колонка `Артикул`, потому что она похожа на `nmID` карточки WB.

Candidate URL строится так:

`https://www.wildberries.ru/catalog/{nm_id}/detail.aspx`

Это именно research-гипотеза первого шага. На следующем шаге мы проверяем её на live-страницах.


## GUI-режим

Добавлен лёгкий desktop GUI на `tkinter` без новых зависимостей.

Запуск из PowerShell / VS Code terminal:

```powershell
wb-inn-gui
```

или так:

```powershell
python -m wb_inn_extractor.gui
```

Что умеет GUI на текущем шаге:

- выбрать входной Excel;
- запустить `analyze`;
- собрать `research_sample.xlsx`;
- запустить `inspect-row` по выбранной строке;
- увидеть лог прямо в окне.

Что важно:

- GUI пока не заменяет будущий batch extractor, а даёт удобный оконный интерфейс для research MVP;
- для `Inspect row` сначала нужно создать `research sample`;
- номер строки для inspect — это номер строки в `research_sample.xlsx`, включая заголовок, поэтому первая строка данных обычно начинается с `2`.


## Новый режим: ручная сессия WB с постоянным профилем

Добавлен режим, который нужен именно для обхода текущего узкого места — антибота Wildberries.

Что делает режим:

- запускает Chromium в `persistent profile` режиме;
- использует отдельную папку профиля, чтобы куки и локальная сессия сохранялись;
- открывает карточку товара;
- даёт время пройти антибот-проверку руками;
- после паузы сохраняет screenshot, HTML, text dump и JSON-результат.

### CLI

```powershell
python -m wb_inn_extractor.cli manual-session --input "C:\path\output
esearch_sample.xlsx" --row 2 --profile-dir "C:\path\wb_profile" --artifacts-dir "C:\pathrtifacts" --wait-seconds 90
```

### GUI

В GUI добавлены:

- поле `Папка профиля WB`;
- поле `Секунд на ручную проверку`;
- кнопка `4. Ручная сессия WB`.

Практический сценарий:

1. создаёшь `research sample`;
2. выбираешь строку;
3. жмёшь `4. Ручная сессия WB`;
4. в открывшемся Chromium, если нужно, проходишь антибот руками;
5. ждёшь завершения таймера;
6. смотришь JSON-статус и артефакты.

### Новые статусы

- `ANTI_BOT_PAGE`
- `MANUAL_CHECK_REQUIRED`
- `PRODUCT_PAGE_OPENED`
- `PAGE_OPENED_NO_REQUISITES`
- `SUCCESS`


## Шаг 5: улучшенный ручной режим

- папка артефактов и папка профиля WB теперь должны быть разными;
- добавлен статус `PARTIAL_SUCCESS`, если часть реквизитов уже найдена, даже когда WB отдал защиту или плохой HTTP-статус;
- GUI теперь явно подсказывает, что делать в ручной сессии;
- добавлены кнопки открытия папки артефактов и папки профиля.

Рекомендуемая схема путей:

- артефакты: `C:\...\тест_парсер\artifacts`
- профиль WB: `C:\...\тест_парсер\wb_profile`


## Step 7

В `inspect-row` добавлен автоматический переход на страницу продавца и попытка раскрыть tooltip с реквизитами.
В JSON теперь пишутся `seller_url` и `navigated_to_seller_page`.


## Step 8

Точный фикс после регрессии:

- обычный `inspect-row` и кнопка `3. Проверить строку` теперь используют `Папка профиля WB`, если она заполнена;
- добавлен более осторожный захват HTML/текста страницы, чтобы не падать на ошибке `Page.content: ... page is navigating`;
- переход на seller page теперь в первую очередь делается через прямой `href` на `/seller/...`, без лишнего клика по карточке.


## Step 11

Точный фикс под seller page:

- реквизиты теперь парсятся сначала из живого tooltip, потом из полного HTML финальной seller page, а не из усечённого `html[:100_000]`;
- это важно, потому что `tooltip-supplier` на WB часто находится далеко внизу HTML и раньше просто не попадал в поиск;
- имя продавца теперь тоже пытаемся брать из tooltip/HTML seller page, а не из title сайта Wildberries.
