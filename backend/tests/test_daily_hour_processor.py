import random

from app.services.daily_hour_processor import (
    HOURS,
    WIDELOAD_DISTRIBUTION_INDEXES,
    WIDELOAD_NORMAL_MAX_PER_HOUR,
    WIDELOAD_PEAK_INDEXES,
    WIDELOAD_PEAK_MAX_PER_HOUR,
    distribute_wideloads,
)


def test_wideload_distribution_stays_inside_allowed_hours():
    wideload_count = 37
    distribution = distribute_wideloads(wideload_count, rng=random.Random(4))

    assert sum(distribution) == wideload_count

    for index, value in enumerate(distribution):
        if index in WIDELOAD_DISTRIBUTION_INDEXES:
            assert value >= 0
        else:
            assert value == 0

    assert HOURS[WIDELOAD_DISTRIBUTION_INDEXES[0]] == "0600-0700"
    assert HOURS[WIDELOAD_DISTRIBUTION_INDEXES[-1]] == "1700-1800"


def test_wideload_distribution_is_lumpy_not_linear():
    distribution = distribute_wideloads(48, rng=random.Random(7))
    allowed_values = [distribution[index] for index in WIDELOAD_DISTRIBUTION_INDEXES]

    assert sum(allowed_values) == 48
    assert len(set(allowed_values)) > 2
    assert max(allowed_values) - min(allowed_values) >= 5


def test_wideload_distribution_caps_high_values_to_peak_windows():
    distribution = distribute_wideloads(72, rng=random.Random(7))
    allowed_values = [distribution[index] for index in WIDELOAD_DISTRIBUTION_INDEXES]
    middle_hour_values = [
        distribution[index]
        for index in WIDELOAD_DISTRIBUTION_INDEXES
        if index not in WIDELOAD_PEAK_INDEXES
    ]

    assert sum(allowed_values) == 72
    assert max(allowed_values) <= WIDELOAD_PEAK_MAX_PER_HOUR
    assert max(middle_hour_values) <= WIDELOAD_NORMAL_MAX_PER_HOUR

    for index, value in enumerate(distribution):
        if value > WIDELOAD_NORMAL_MAX_PER_HOUR:
            assert index in WIDELOAD_PEAK_INDEXES
