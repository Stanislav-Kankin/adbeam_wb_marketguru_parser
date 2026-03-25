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
