from datetime import datetime

from core.sla import business_hours_between


def dt(day, hour):
    return datetime(2024, 1, day, hour, 0, 0)


def test_business_hours_simple_daytime():
    assert business_hours_between(dt(1, 10), dt(1, 18), start_hour=9, end_hour=18) == 8


def test_business_hours_weekend_skip():
    # Jan 6-7 2024 were weekend; should be zero
    assert business_hours_between(dt(6, 10), dt(7, 18)) == 0


def test_business_hours_holiday():
    h = ["2024-01-02"]
    assert business_hours_between(dt(2, 9), dt(2, 18), holidays=h) == 0


def test_business_hours_cross_midnight():
    # From 18:00 to next day 11:00 with 9-18 hours => only 2 hours (next day 9-11)
    assert business_hours_between(dt(1, 18), dt(2, 11), start_hour=9, end_hour=18) == 2
