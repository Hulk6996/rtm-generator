#!/usr/bin/env python3
"""
main.py — точка входа RTM Generator.

Использование:
  python main.py                           # использует INPUT_FILE из config.py
  python main.py data/my_requirements.xlsx # кастомный файл
"""

import sys
import os
import time

# Добавляем src в путь
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from config import INPUT_FILE, EXCEL_OUT, WORD_OUT
from src.parser     import load_requirements
from src.rtm_builder import build_rtm
from src.analyzer   import analyze
from src.excel_report import generate_excel
from src.word_report  import generate_word


def main(input_file: str | None = None):
    filepath = input_file or INPUT_FILE
    t0 = time.time()

    print("=" * 56)
    print("  RTM Generator — Матрица трассировки требований")
    print("=" * 56)
    print(f"\n📂 Входной файл: {filepath}")
    print()

    # 1. Загрузка данных
    print("📥 Загрузка требований...")
    data = load_requirements(filepath)

    # 2. Построение RTM
    print("🔗 Построение матрицы трассировки...")
    rtm = build_rtm(data)
    print(f"  ✔ Строк RTM: {len(rtm['rtm_rows'])}")

    # 3. Анализ
    print("📊 Анализ покрытия и метрик...")
    metrics = analyze(rtm, data)

    _print_summary(metrics)

    # 4. Excel-отчёт
    print("\n📝 Генерация Excel-отчёта...")
    generate_excel(rtm, metrics, EXCEL_OUT)

    # 5. Word-отчёт
    print("📄 Генерация Word-резюме...")
    generate_word(rtm, metrics, WORD_OUT)

    elapsed = time.time() - t0
    print(f"\n✅ Готово за {elapsed:.1f} с")
    print(f"   Excel: {EXCEL_OUT}")
    print(f"   Word:  {WORD_OUT}")
    print("=" * 56)


def _print_summary(metrics: dict):
    br  = metrics["br_coverage"]
    fr  = metrics["fr_coverage"]
    ts  = metrics["test_stats"]
    h   = metrics["health"]

    print()
    print("┌────────────────────────────────────────────┐")
    print("│              СВОДКА МЕТРИК                  │")
    print("├──────────────────────────┬──────────────────┤")
    print(f"│  BR → FR покрытие        │  {br['pct_fr']:>5.1f}%          │")
    print(f"│  FR → TC покрытие        │  {fr['pct']:>5.1f}%          │")
    print(f"│  Тесты Passed            │  {ts['pct_pass']:>5.1f}%          │")
    print(f"│  Health Score            │  {h['score']:>5.1f}/100       │")
    print(f"│  Статус                  │  {h['label']:<16} │")
    print("├──────────────────────────┴──────────────────┤")
    print(f"│  Непокрытых BR: {len(metrics['uncovered_br']):<3}  Непокрытых FR: {len(metrics['uncovered_fr']):<3}  │")
    if metrics["duplicates"]:
        print(f"│  ⚠  Дублей: {len(metrics['duplicates']):<3}                              │")
    if metrics["conflicts"]:
        print(f"│  ⚡ Конфликтов: {len(metrics['conflicts']):<3}                           │")
    print("└─────────────────────────────────────────────┘")


if __name__ == "__main__":
    file_arg = sys.argv[1] if len(sys.argv) > 1 else None
    main(file_arg)
