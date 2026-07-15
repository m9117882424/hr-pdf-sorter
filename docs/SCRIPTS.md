# Описание скриптов

Документ фиксирует назначение, входные данные, выходные данные и порядок запуска основных утилит репозитория.

## Краткая карта

| Скрипт | Для чего нужен | Основной запуск |
| --- | --- | --- |
| `sort_payment_pdfs.py` | Раскладка PDF по папкам сотрудников по ФИО в имени файла. | `py sort_payment_pdfs.py --dir "папка" --move` |
| `ru_decont.py` | Базовый движок сверки PDF-платёжек Ziraat с Excel. Поддерживает русские и турецкие формы. | `py ru_decont.py --excel "input.xlsx" --pdf-dir "pp_ru"` |
| `ru_decont_fast.py` | Быстрый запускатель сверки: PDF обрабатываются параллельно, Excel заполняется последовательно. | `py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "pp_ru" --workers 3` |
| `compare_ru_decont.bat` | Windows-запуск ускоренного режима. | двойной клик по BAT |

---

## `sort_payment_pdfs.py`

### Назначение

Раскладывает PDF-файлы платёжных документов по папкам сотрудников. Основной источник ФИО — имя PDF-файла.

Пример имени:

```text
ПП от 02.10.2025 Иванов Иван_rus.pdf
```

### Как работает

1. Ищет PDF-файлы в рабочей папке.
2. Ищет папки сотрудников.
3. Чистит имя файла от служебных слов: `ПП`, `от`, `rus`, `pdf`, даты и похожего мусора.
4. Нормализует ФИО и названия папок.
5. Считает совпадение через токены и fuzzy-score.
6. В безопасном режиме показывает план перемещения.
7. В боевом режиме перемещает файл в найденную папку.

### Основные команды

Проверка без перемещения:

```bat
py sort_payment_pdfs.py --dir "C:\Users\Maksim\Desktop\для HR"
```

Реальное перемещение:

```bat
py sort_payment_pdfs.py --dir "C:\Users\Maksim\Desktop\для HR" --move
```

Отчёт:

```bat
py sort_payment_pdfs.py --report reports\sort_report.csv
```

### Важные параметры

| Параметр | Назначение |
| --- | --- |
| `--dir` | Рабочая папка. По умолчанию текущая. |
| `--move` | Включает реальное перемещение. Без него только проверка. |
| `--min-score` | Минимальный порог совпадения. По умолчанию `0.78`. |
| `--ambiguity-gap` | Если лучший и второй кандидат слишком близки, файл пропускается. |
| `--recursive-pdfs` | Искать PDF рекурсивно. |
| `--recursive-folders` | Искать папки сотрудников рекурсивно. |
| `--report` | Путь к CSV-отчёту. |

---

## `ru_decont.py`

### Назначение

Базовый движок сверки PDF-платёжек Ziraat Bankasi с Excel-реестром. Скрипт извлекает из PDF дату, ИНН/кимлик, ФИО и суммы, затем дописывает данные в найденные строки Excel.

Поддерживаются два типа документов:

1. Русская форма: `ИНН / Кимлик`, `Фамилия / Наименование`, `Имя`, суммы в русскоязычной таблице.
2. Турецкая форма: `Vergi Kimlik Numarası`, `TCKNO/VKN:...`, `Soyadı / Ünvanı`, `Adı`, `Tarih`, суммы в столбце `Miktarı`.

### Как работает

1. Открывает Excel.
2. Находит строку заголовков с обязательными колонками.
3. Строит индексы по ФИО, дате и кимлику.
4. Собирает PDF из указанной папки.
5. Для текстовых PDF сначала использует `pdfplumber`.
6. Для сканов и проблемных PDF использует OCR через Tesseract.
7. Для многостраничных PDF определяет полезные страницы с таблицами и пропускает почти пустые страницы.
8. Сопоставляет PDF со строкой Excel.
9. Записывает результат в колонки начиная с `AH`.
10. Создаёт итоговый Excel и CSV-отчёты.

### Когда использовать

Использовать для диагностики или одиночных прогонов, когда скорость не критична. Для рабочих пачек лучше использовать `ru_decont_fast.py`.

### Обязательные колонки Excel

```text
Y.T.C № Кимлики
Дата оплаты
ФИО или ФИО рус
```

Для турецких платёжек латинское имя/фамилия берётся из колонки `ФИО`.

Пример:

```text
ФИО: VASILII SHIPULIN
PDF: Soyadı / Ünvanı : SHIPULIN
PDF: Adı : VASILII
```

Скрипт проверяет оба порядка имени:

```text
SHIPULIN VASILII
VASILII SHIPULIN
```

### Запуск

```bat
py ru_decont.py --excel "input.xlsx" --pdf-dir "pp_ru"
```

Если Excel лежит в подпапке:

```bat
py ru_decont.py --excel "excel_input\input.xlsx" --pdf-dir "pp_ru"
```

С конкретным листом:

```bat
py ru_decont.py --excel "input.xlsx" --pdf-dir "pp_ru" --sheet "Лист1"
```

Без OCR:

```bat
py ru_decont.py --excel "input.xlsx" --pdf-dir "pp_ru" --no-ocr
```

---

## `ru_decont_fast.py`

### Назначение

Ускоренный запускатель для сверки PDF с Excel. Он использует функции из `ru_decont.py`, но обрабатывает PDF параллельно в нескольких процессах.

Excel заполняется последовательно после завершения OCR/парсинга, поэтому гонок при записи в Excel нет.

### Важно

`ru_decont_fast.py` не заменяет `ru_decont.py`. Он импортирует `ru_decont.py` как основной модуль.

Оба файла должны лежать рядом:

```text
для HR/
├─ ru_decont.py
├─ ru_decont_fast.py
├─ input.xlsx
└─ pp_ru/
   ├─ file1.pdf
   └─ file2.pdf
```

Если рядом нет `ru_decont.py`, будет ошибка:

```text
ModuleNotFoundError: No module named 'ru_decont'
```

Если папка `pp_ru` пустая или PDF лежат в другом месте, будет ошибка:

```text
FileNotFoundError: No PDF files in folder: pp_ru
```

### Быстрая проверка перед запуском

```bat
dir ru_decont.py ru_decont_fast.py compare_ru_decont.bat input.xlsx
dir pp_ru\*.pdf
```

### Основной запуск

```bat
py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "pp_ru" --workers 3
```

Если Excel лежит в подпапке:

```bat
py ru_decont_fast.py --excel "excel_input\input.xlsx" --pdf-dir "pp_ru" --workers 3
```

Если PDF лежат рядом со скриптом, а не в `pp_ru`:

```bat
py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "." --workers 3
```

С конкретным листом:

```bat
py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "pp_ru" --sheet "Лист1" --workers 3
```

Без OCR:

```bat
py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "pp_ru" --no-ocr --workers 3
```

### Настройка скорости

```bat
py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "pp_ru" --workers 2
py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "pp_ru" --workers 3
py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "pp_ru" --workers 4
```

Рекомендации:

- `--workers 2` — для слабого ПК или ноутбука;
- `--workers 3` — рекомендуемый баланс;
- `--workers 4` — для более мощной машины;
- больше `4` обычно не нужно, потому что Tesseract сильно грузит CPU и память.

---

## `compare_ru_decont.bat`

### Назначение

Упрощённый Windows-запуск ускоренной сверки.

Текущая команда:

```bat
py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "pp_ru" --workers 3
```

### Когда править BAT

Править, если:

- Excel лежит в другом месте;
- папка PDF называется иначе;
- PDF лежат рядом со скриптом, а не в `pp_ru`;
- нужно указать конкретный лист Excel;
- нужно временно отключить OCR;
- нужно изменить количество `--workers`.

---

## Колонки записи в Excel

| Колонка | Поле |
| --- | --- |
| `AH` | Дата |
| `AI` | ИНН / Кимлик |
| `AJ` | Фамилия / Наименование |
| `AK` | Имя |
| `AL` | Сумма |
| `AM` | Сумма 2 |

---

## Логика сопоставления PDF с Excel

Приоритеты:

1. `ФИО + дата + кимлик`;
2. `дата + кимлик`, если в Excel такая строка одна;
3. `ФИО + дата`, если найдено одно совпадение;
4. `ФИО + кимлик`, если найдено одно совпадение;
5. только `ФИО`, если сотрудник в Excel встречается один раз.

Для латиницы скрипт учитывает оба порядка ФИО: `фамилия имя` из PDF и `имя фамилия` из Excel.

Если совпадение неоднозначное или отсутствует, Excel не заполняется. Запись уходит в `RU_decont_unmatched_report.csv`.

---

## OCR-особенности

Скрипт умеет:

- читать сканированные PDF;
- находить дату в поле `Tarih` и в русских полях даты;
- находить кимлик в `TCKNO/VKN:...`, `Vergi Kimlik Numarası`, `ИНН / Кимлик`;
- находить ФИО в `Soyadı / Ünvanı`, `Adı`, `Фамилия / Наименование`, `Фамилия / Должность`, `Фамилия / Звание`;
- извлекать суммы из столбца `Miktarı` и русскоязычных таблиц;
- пропускать почти пустые страницы в многостраничных PDF;
- фиксировать полезные и пропущенные страницы в отчётах.

---

## Диагностика

### PDF не сопоставился

Открыть `RU_decont_unmatched_report.csv` и проверить:

- `Дата`;
- `ИНН / Кимлик`;
- `Фамилия / Наименование`;
- `Имя`;
- `match_key_name`;
- `match_key_name_reversed`;
- `match_key_date`;
- `match_key_id`;
- `match_mode`.

Если ФИО пустое, но дата и кимлик есть, скрипт попробует найти строку по `дата + кимлик`. Если всё равно `unmatched`, значит в Excel нет уникальной строки с такой парой.

### OCR подозрительный

Открыть `RU_decont_ocr_debug_report.csv` и проверить:

- `raw_top_date_ocr`;
- `raw_kimlik_ocr`;
- `raw_name_ocr`;
- `payment_pages`;
- `skipped_pages`.

### Нет PDF в папке

```bat
dir pp_ru\*.pdf
```

Если PDF лежат рядом со скриптом:

```bat
py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "." --workers 3
```

---

## Рекомендуемый рабочий процесс

1. Сложить PDF в `pp_ru` или явно указать свою папку через `--pdf-dir`.
2. Проверить наличие файлов командами `dir`.
3. Запустить `compare_ru_decont.bat` или `ru_decont_fast.py`.
4. Проверить `RU_decont_unmatched_report.csv`.
5. Проверить `RU_decont_ocr_debug_report.csv`, если есть проблемы с OCR.
6. Использовать итоговый `input_RU_decont_filled.xlsx`.

Код должен ехать в Git, персональные документы — нет.
