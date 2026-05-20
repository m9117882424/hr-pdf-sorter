# HR PDF Sorter

Набор утилит для обработки HR-платежек в PDF:

1. раскладка PDF по папкам сотрудников;
2. сверка русских платежек с Excel;
3. OCR-распознавание сканов и многостраничных документов;
4. ускоренная параллельная обработка пачки PDF.

## Состав проекта

| Файл | Назначение |
| --- | --- |
| `sort_payment_pdfs.py` | Раскладывает PDF-файлы по папкам сотрудников по ФИО в имени файла. |
| `ru_decont.py` | Базовый движок сверки PDF с Excel. Извлекает дату, ИНН/кимлик, ФИО и суммы, затем заполняет Excel начиная с `AH`. |
| `ru_decont_fast.py` | Быстрый запускатель. Обрабатывает PDF параллельно, Excel заполняет последовательно. |
| `compare_ru_decont.bat` | Windows-запуск ускоренного режима. |
| `requirements.txt` | Python-зависимости. |
| `.gitignore` | Исключения для рабочих файлов и отчетов. |

Подробное описание скриптов: [`docs/SCRIPTS.md`](docs/SCRIPTS.md).

## Установка

Требуется Python 3.10+.

```bat
py -m pip install -r requirements.txt
```

Для OCR нужен Tesseract OCR. На Windows ожидается стандартный путь:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

Для русских сканов желательно установить языковой пакет `rus.traineddata`.

## Сценарий 1. Раскладка PDF по папкам сотрудников

Проверка без перемещения:

```bat
py sort_payment_pdfs.py --dir "C:\Users\Maksim\Desktop\для HR"
```

Реальное перемещение:

```bat
py sort_payment_pdfs.py --dir "C:\Users\Maksim\Desktop\для HR" --move
```

Отчет:

```bat
py sort_payment_pdfs.py --report reports\sort_report.csv
```

## Сценарий 2. Сверка русских PDF-платежек с Excel

### Ожидаемая структура для быстрого запуска

Если используется `compare_ru_decont.bat`, рядом с BAT должны лежать оба Python-файла, Excel и папка `pp_ru` с PDF:

```text
для HR/
├─ ru_decont.py
├─ ru_decont_fast.py
├─ compare_ru_decont.bat
├─ input.xlsx
└─ pp_ru/
   ├─ file1.pdf
   └─ file2.pdf
```

Важно: `ru_decont_fast.py` не заменяет `ru_decont.py`. Он импортирует его как основной модуль, поэтому оба файла должны быть в одной папке.

Если рядом нет `ru_decont.py`, будет ошибка:

```text
ModuleNotFoundError: No module named 'ru_decont'
```

Если в `pp_ru` нет PDF, будет ошибка:

```text
FileNotFoundError: No PDF files in folder: pp_ru
```

### Быстрая проверка перед запуском

В папке проекта выполнить:

```bat
dir ru_decont.py ru_decont_fast.py compare_ru_decont.bat input.xlsx
dir pp_ru\*.pdf
```

Должны быть найдены:

```text
ru_decont.py
ru_decont_fast.py
compare_ru_decont.bat
input.xlsx
минимум один PDF в pp_ru
```

### Обязательные колонки в Excel

```text
ФИО рус
Дата оплаты
Y.T.C № Кимлики
```

### Запуск через BAT

```bat
compare_ru_decont.bat
```

Текущий BAT запускает ускоренный режим:

```bat
py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "pp_ru" --workers 3
```

### Запуск вручную

Быстрый режим:

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

Базовый режим для диагностики:

```bat
py ru_decont.py --excel "input.xlsx" --pdf-dir "pp_ru"
```

### Настройка скорости

`--workers` задает количество параллельных процессов OCR/парсинга PDF:

```bat
py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "pp_ru" --workers 2
py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "pp_ru" --workers 3
py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "pp_ru" --workers 4
```

Рекомендации:

- `--workers 2` — для слабого ПК;
- `--workers 3` — рекомендуемый баланс;
- `--workers 4` — для более мощной машины;
- больше `4` обычно не нужно.

## Что заполняет сверка

Данные записываются в исходный Excel начиная с колонки `AH`:

| Колонка | Поле |
| --- | --- |
| `AH` | Дата |
| `AI` | ИНН / Кимлик |
| `AJ` | Фамилия / Наименование |
| `AK` | Имя |
| `AL` | Сумма |
| `AM` | Сумма 2 |

Итоговый файл:

```text
input_RU_decont_filled.xlsx
```

Отчеты:

```text
RU_decont_parsed_report.csv
RU_decont_unmatched_report.csv
RU_decont_ocr_debug_report.csv
```

## Логика сопоставления PDF с Excel

1. `ФИО + дата + кимлик`;
2. `дата + кимлик`, если в Excel такая строка одна;
3. `ФИО + дата`, если совпадение единственное;
4. `ФИО + кимлик`, если совпадение единственное;
5. только `ФИО`, если человек в Excel встречается один раз.

Второй пункт нужен для сканов, где OCR плохо читает ФИО или имя файла повреждено кодировкой, но дата и кимлик распознаны корректно.

## Типовые проблемы

### `ModuleNotFoundError: No module named 'ru_decont'`

Рядом с `ru_decont_fast.py` нет файла `ru_decont.py`. Решение:

```bat
git pull origin main
```

Или скопировать `ru_decont.py` рядом с `ru_decont_fast.py`.

### `FileNotFoundError: No PDF files in folder: pp_ru`

Папка `pp_ru` пустая, отсутствует или PDF лежат в другом месте. Проверка:

```bat
dir pp_ru\*.pdf
```

Если PDF лежат рядом со скриптом:

```bat
py ru_decont_fast.py --excel "input.xlsx" --pdf-dir "." --workers 3
```

### PDF попал в `unmatched`

Проверить `RU_decont_unmatched_report.csv`:

- распознанное ФИО;
- дату;
- кимлик;
- `match_mode`.

Если ФИО пустое, но дата и кимлик есть, скрипт попробует найти строку по `дата + кимлик`. Если все равно `unmatched`, значит в Excel нет уникальной строки с такой парой.

### OCR не видит ФИО

Проверить `RU_decont_ocr_debug_report.csv`, поле `raw_name_ocr`.

### Tesseract не найден

Проверить:

```bat
"C:\Program Files\Tesseract-OCR\tesseract.exe" --version
```

## Безопасность данных

Не добавлять в Git реальные PDF, Excel, CSV-отчеты, скриншоты документов и рабочие файлы с персональными данными.

Перед коммитом проверять:

```bat
git status
```
