"""
tests/test_parser.py — базовые тесты парсера и анализатора.
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from src.rtm_builder import build_rtm
from src.analyzer    import analyze


# ── Мок-данные ───────────────────────────────────────────────

MOCK_DATA = {
    "br": [
        {"id": "BR-001", "title": "Авторизация", "description": "Система должна обеспечивать авторизацию",
         "priority": "High", "category": "Auth", "source": "Аналитик", "status": "Active"},
        {"id": "BR-002", "title": "Отчёты", "description": "Пользователь должен иметь доступ к отчётам",
         "priority": "Medium", "category": "Reporting", "source": "Аналитик", "status": "Active"},
        {"id": "BR-003", "title": "Уведомления", "description": "Необходимо обеспечить уведомления",
         "priority": "Low", "category": "Notifications", "source": "Аналитик", "status": "Active"},
    ],
    "fr": [
        {"id": "FR-001", "title": "Логин по паролю", "description": "...",
         "br_refs": ["BR-001"], "type": "Functional", "component": "Backend", "status": "Active"},
        {"id": "FR-002", "title": "JWT-токен", "description": "...",
         "br_refs": ["BR-001"], "type": "Security", "component": "Auth Service", "status": "Active"},
        {"id": "FR-003", "title": "Генерация Excel", "description": "...",
         "br_refs": ["BR-002"], "type": "Functional", "component": "Backend", "status": "Active"},
        {"id": "FR-004", "title": "Orphan FR", "description": "...",
         "br_refs": [], "type": "Functional", "component": "Backend", "status": "Draft"},
    ],
    "tc": [
        {"id": "TC-001", "title": "Тест логина", "fr_refs": ["FR-001"],
         "type": "Manual", "result": "Passed", "priority": "High"},
        {"id": "TC-002", "title": "Тест токена", "fr_refs": ["FR-001", "FR-002"],
         "type": "Auto",   "result": "Passed", "priority": "High"},
        {"id": "TC-003", "title": "Тест отчёта", "fr_refs": ["FR-003"],
         "type": "Manual", "result": "Failed", "priority": "Medium"},
        {"id": "TC-004", "title": "Orphan TC",   "fr_refs": [],
         "type": "Manual", "result": "Not Run",  "priority": "Low"},
    ],
}


# ── Tests ─────────────────────────────────────────────────────

def test_build_rtm_links():
    rtm = build_rtm(MOCK_DATA)
    assert rtm["br_to_frs"]["BR-001"] == ["FR-001", "FR-002"]
    assert rtm["br_to_frs"]["BR-002"] == ["FR-003"]
    assert rtm["br_to_frs"]["BR-003"] == []
    print("✅ test_build_rtm_links passed")


def test_orphan_detection():
    rtm = build_rtm(MOCK_DATA)
    assert "FR-004" in rtm["orphan_fr"]
    assert "TC-004" in rtm["orphan_tc"]
    print("✅ test_orphan_detection passed")


def test_rtm_rows():
    rtm = build_rtm(MOCK_DATA)
    rows = rtm["rtm_rows"]
    assert len(rows) > 0
    # BR-003 should appear once (no FR)
    br3_rows = [r for r in rows if r["br_id"] == "BR-003"]
    assert len(br3_rows) == 1
    assert br3_rows[0]["fr_id"] == ""
    print("✅ test_rtm_rows passed")


def test_coverage_metrics():
    rtm     = build_rtm(MOCK_DATA)
    metrics = analyze(rtm, MOCK_DATA)

    # BR-001 and BR-002 covered, BR-003 not
    assert metrics["br_coverage"]["with_fr"] == 2
    assert metrics["br_coverage"]["pct_fr"]  == round(2/3*100, 1)

    # FR-001, FR-002, FR-003 covered by TC; FR-004 not
    assert metrics["fr_coverage"]["with_tc"] == 3
    print("✅ test_coverage_metrics passed")


def test_test_stats():
    rtm     = build_rtm(MOCK_DATA)
    metrics = analyze(rtm, MOCK_DATA)
    ts = metrics["test_stats"]
    assert ts["total"]   == 4
    assert ts["passed"]  == 2
    assert ts["failed"]  == 1
    assert ts["not_run"] == 1
    print("✅ test_test_stats passed")


def test_uncovered_items():
    rtm     = build_rtm(MOCK_DATA)
    metrics = analyze(rtm, MOCK_DATA)
    assert "BR-003" in metrics["uncovered_br"]
    assert "FR-004" in metrics["uncovered_fr"]
    print("✅ test_uncovered_items passed")


def test_health_score():
    rtm     = build_rtm(MOCK_DATA)
    metrics = analyze(rtm, MOCK_DATA)
    score = metrics["health"]["score"]
    assert 0 <= score <= 100
    print(f"✅ test_health_score passed  (score={score})")


# ── Runner ────────────────────────────────────────────────────

if __name__ == "__main__":
    tests = [
        test_build_rtm_links,
        test_orphan_detection,
        test_rtm_rows,
        test_coverage_metrics,
        test_test_stats,
        test_uncovered_items,
        test_health_score,
    ]
    failed = 0
    print(f"\n{'='*46}")
    print("  RTM Generator — Test Suite")
    print(f"{'='*46}\n")
    for t in tests:
        try:
            t()
        except AssertionError as e:
            print(f"❌ {t.__name__} FAILED: {e}")
            failed += 1
        except Exception as e:
            print(f"💥 {t.__name__} ERROR: {e}")
            failed += 1

    print(f"\n{'─'*46}")
    if failed == 0:
        print(f"  ✅ Все {len(tests)} тестов прошли успешно!")
    else:
        print(f"  ❌ Провалено: {failed} из {len(tests)}")
    print(f"{'='*46}\n")
