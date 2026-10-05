from datetime import date, datetime

from app import Entry, entries_for, period_bounds, week_start_for


def entry(day: str) -> Entry:
    value = datetime.fromisoformat(f"{day}T12:00:00")
    return Entry(1, value, "Title", "Description")


def test_week_starts_on_monday():
    assert week_start_for(date(2026, 10, 7)) == date(2026, 10, 5)


def test_daily_bounds():
    assert period_bounds("Daily", date(2026, 10, 5)) == (date(2026, 10, 5), date(2026, 10, 5))


def test_weekly_filter_includes_monday_through_sunday():
    values = [entry("2026-10-05"), entry("2026-10-11"), entry("2026-10-12")]
    result = entries_for(values, "Weekly", date(2026, 10, 7))
    assert [item.completed_at.date() for item in result] == [date(2026, 10, 5), date(2026, 10, 11)]


def test_monthly_filter_handles_month_end():
    values = [entry("2026-02-28"), entry("2026-03-01")]
    result = entries_for(values, "Monthly", date(2026, 2, 1))
    assert len(result) == 1
    assert result[0].completed_at.date() == date(2026, 2, 28)
