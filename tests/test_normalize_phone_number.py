import sys
from pathlib import Path

import pytest

sys.path.append(str(Path(__file__).resolve().parents[1]))
from synchronoss_parser.utils import normalize_phone_number


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("1112223333", "1112223333"),
        ("+1 111-222-3333", "1112223333"),
        ("(111) 222-3333", "1112223333"),
        ("  +1-111-222-3333  ", "1112223333"),
        ("+12223334444", "2223334444"),
        ("12223334444", "2223334444"),
    ],
)
def test_normalize_phone_number(raw, expected):
    assert normalize_phone_number(raw) == expected
