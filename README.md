# 📋 RTM Generator — Генератор Матрицы Трассировки Требований

> **Инструмент системного аналитика** для автоматической генерации Requirements Traceability Matrix (RTM) из Excel-файла с требованиями.
> Результат — профессиональный Excel-отчёт (4 листа) + Word-резюме с KPI, анализом пробелов и рекомендациями.

---

## 🚀 Возможности

| Функция | Описание |
|---|---|
| 🔗 **Матрица трассировки** | Связывает BR → FR → TC в единую таблицу |
| 📊 **Покрытие требований** | % BR покрытых функциональными требованиями и тестами |
| 🧪 **Статистика тестов** | Passed / Failed / Blocked / Not Run с цветовой индикацией |
| 📈 **Health Score** | Интегральный балл качества матрицы (0–100) |
| 🔍 **Анализ по приоритетам** | Покрытие High / Medium / Low требований |
| ⚠ **Обнаружение проблем** | Дубли, конфликты, «сироты», непокрытые требования |
| 📝 **Excel + Word отчёт** | Стилизованный Excel (4 листа) и Word с резюме |

---

## 📂 Структура проекта

```
rtm-generator/
├── main.py                    # Точка входа
├── config.py                  # Конфигурация (пути, цвета, метаданные)
├── requirements.txt
│
├── src/
│   ├── parser.py              # Чтение Excel (BR / FR / TC)
│   ├── rtm_builder.py         # Построение матрицы BR → FR → TC
│   ├── analyzer.py            # Метрики, покрытие, дубли, конфликты
│   ├── excel_report.py        # Генерация Excel-отчёта (4 листа)
│   └── word_report.py         # Генерация Word-резюме
│
├── data/
│   ├── requirements_sample.xlsx   # Демо входной файл (30 BR / 46 FR / 61 TC)
│   └── generate_sample.py         # Скрипт генерации демо-данных
│
├── output/                    # Сюда сохраняются отчёты
│   ├── rtm_report.xlsx
│   └── rtm_summary.docx
│
└── tests/
    └── test_parser.py         # 7 unit-тестов
```

---

## ⚙️ Установка и запуск

### 1. Клонирование

```bash
git clone https://github.com/Hulk6996/rtm-generator.git
cd rtm-generator
```

### 2. Зависимости

```bash
pip install -r requirements.txt
```

### 3. Запуск с демо-данными

```bash
python main.py
```

Или с собственным файлом:

```bash
python main.py путь/к/requirements.xlsx
```

### 4. Тесты

```bash
python tests/test_parser.py
```

---

## 📥 Формат входного Excel

Файл должен содержать **3 листа**:

### Лист 1: `Business Requirements`

| BR_ID | Title | Description | Priority | Category | Source | Status |
|---|---|---|---|---|---|---|
| BR-001 | Авторизация | Система должна... | High | Auth | Аналитик | Active |

### Лист 2: `Functional Requirements`

| FR_ID | Title | Description | BR_REF | Type | Component | Status |
|---|---|---|---|---|---|---|
| FR-001 | Логин по паролю | ... | BR-001 | Functional | Backend | Active |
| FR-002 | JWT-токен | ... | BR-001, BR-002 | Security | Auth Service | Active |

> `BR_REF` — через запятую, например: `BR-001, BR-003`

### Лист 3: `Test Cases`

| TC_ID | Title | FR_REF | Type | Result | Priority |
|---|---|---|---|---|---|
| TC-001 | Тест логина | FR-001 | Manual | Passed | High |
| TC-002 | Тест токена | FR-001, FR-002 | Auto | Failed | High |

> `Result`: `Passed` / `Failed` / `Blocked` / `Not Run`

---

## 📊 Выходные отчёты

### Excel (`output/rtm_report.xlsx`)

| Лист | Содержимое |
|---|---|
| 📊 Дашборд | KPI-плитки, покрытие по приоритетам, результаты тестов |
| 📋 RTM | Полная матрица BR → FR → TC с цветовым кодированием |
| 🔍 Покрытие BR | Таблица покрытия по каждому BR с индикаторами |
| ⚠ Проблемы | Непокрытый BR/FR, сироты, дубли, конфликты |

### Word (`output/rtm_summary.docx`)

- Титульный лист с датой
- Сводка KPI (таблица)
- Покрытие по приоритетам
- Полная RTM-таблица
- Список проблем
- Рекомендации по устранению

---

## 📈 Метрики и Health Score

```
Health Score = BR_coverage × 0.40 + FR_coverage × 0.30 + Test_pass_rate × 0.30
```

| Диапазон | Статус |
|---|---|
| ≥ 80 | 🟢 Хорошее |
| 50–79 | 🟡 Среднее |
| < 50 | 🔴 Критическое |

---

## 🔧 Конфигурация

Все параметры в `config.py`:

```python
INPUT_FILE  = "data/requirements_sample.xlsx"
EXCEL_OUT   = "output/rtm_report.xlsx"
WORD_OUT    = "output/rtm_summary.docx"

REPORT_PROJECT = "Мой Проект"
REPORT_VERSION = "2.0"

COVERAGE_GREEN  = 80   # порог зелёного статуса
COVERAGE_YELLOW = 50   # порог жёлтого статуса
```

---

## 🛠 Стек технологий

| Библиотека | Назначение |
|---|---|
| `pandas` | Чтение и обработка Excel |
| `openpyxl` | Генерация стилизованного Excel |
| `python-docx` | Генерация Word-документа |

---

## 📌 Пример вывода

```
========================================================
  RTM Generator — Матрица трассировки трассирований
========================================================

📥 Загрузка требований...
  ✔ BR:  30 | FR:  46 | TC:  61
🔗 Построение матрицы трассировки...
  ✔ Строк RTM: 159

┌─────────────────────────────────────────────┐
│  BR → FR покрытие  │  90.0%                 │
│  FR → TC покрытие  │  87.0%                 │
│  Тесты Passed      │  47.5%                 │
│  Health Score      │  76.3/100              │
│  Статус            │  🟡 Среднее             │
└─────────────────────────────────────────────┘

  ✔ Excel сохранён: output/rtm_report.xlsx
  ✔ Word  сохранён: output/rtm_summary.docx
✅ Готово за 2.8 с
```

---

## 📄 Лицензия

MIT — используйте свободно в коммерческихД4/личных проектахм.