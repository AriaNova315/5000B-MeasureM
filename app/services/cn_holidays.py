from datetime import date

# 法定放假日期（含调休后的实际休息日）
_HOLIDAYS: frozenset[date] = frozenset({
    # 2024
    date(2024, 1, 1),
    date(2024, 2, 10), date(2024, 2, 11), date(2024, 2, 12),
    date(2024, 2, 13), date(2024, 2, 14), date(2024, 2, 15),
    date(2024, 2, 16), date(2024, 2, 17),
    date(2024, 4, 4), date(2024, 4, 5), date(2024, 4, 6),
    date(2024, 5, 1), date(2024, 5, 2), date(2024, 5, 3),
    date(2024, 6, 10),
    date(2024, 9, 15), date(2024, 9, 16), date(2024, 9, 17),
    date(2024, 10, 1), date(2024, 10, 2), date(2024, 10, 3),
    date(2024, 10, 4), date(2024, 10, 5), date(2024, 10, 6), date(2024, 10, 7),
    # 2025
    date(2025, 1, 1),
    date(2025, 1, 28), date(2025, 1, 29), date(2025, 1, 30),
    date(2025, 1, 31), date(2025, 2, 1), date(2025, 2, 2),
    date(2025, 2, 3), date(2025, 2, 4),
    date(2025, 4, 4), date(2025, 4, 5), date(2025, 4, 6),
    date(2025, 5, 1), date(2025, 5, 2), date(2025, 5, 3),
    date(2025, 5, 4), date(2025, 5, 5),
    date(2025, 5, 31), date(2025, 6, 1), date(2025, 6, 2),
    date(2025, 10, 1), date(2025, 10, 2), date(2025, 10, 3),
    date(2025, 10, 4), date(2025, 10, 5), date(2025, 10, 6),
    date(2025, 10, 7), date(2025, 10, 8),
    # 2026
    date(2026, 1, 1), date(2026, 1, 2),
    date(2026, 2, 17), date(2026, 2, 18), date(2026, 2, 19),
    date(2026, 2, 20), date(2026, 2, 21), date(2026, 2, 22),
    date(2026, 2, 23), date(2026, 2, 24),
    date(2026, 4, 4), date(2026, 4, 5), date(2026, 4, 6),
    date(2026, 5, 1), date(2026, 5, 2), date(2026, 5, 3),
    date(2026, 5, 4), date(2026, 5, 5),
    date(2026, 6, 19), date(2026, 6, 20), date(2026, 6, 21),
    date(2026, 10, 1), date(2026, 10, 2), date(2026, 10, 3),
    date(2026, 10, 4), date(2026, 10, 5), date(2026, 10, 6), date(2026, 10, 7),
})

# 调休上班日（本为周末，但须上班）
_MAKEUP_WORKDAYS: frozenset[date] = frozenset({
    # 2024
    date(2024, 2, 4), date(2024, 2, 18),
    date(2024, 4, 7), date(2024, 4, 28),
    date(2024, 5, 11),
    date(2024, 9, 14), date(2024, 9, 29),
    date(2024, 10, 12),
    # 2025
    date(2025, 1, 26), date(2025, 2, 8),
    date(2025, 4, 27),
    date(2025, 9, 28), date(2025, 10, 11),
    # 2026
    date(2026, 1, 4),
    date(2026, 2, 15), date(2026, 2, 28),
    date(2026, 4, 12),
    date(2026, 5, 9),
    date(2026, 9, 27), date(2026, 10, 10),
})


def is_workday(d: date) -> bool:
    if d in _MAKEUP_WORKDAYS:
        return True
    if d in _HOLIDAYS:
        return False
    return d.weekday() < 5
