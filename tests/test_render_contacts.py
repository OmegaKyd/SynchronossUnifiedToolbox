import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.append(str(Path(__file__).resolve().parents[1]))
from synchronoss_parser.render_transcripts import (
    Message,
    build_contact_lookup,
    normalize_phone_number,
    render_thread_html,
)


def test_contact_lookup_replaces_number_with_name(tmp_path):
    df = pd.DataFrame(
        [{"firstname": "Alice", "lastname": "Smith", "phone_numbers": "(123) 456-7890"}]
    )
    xlsx_path = tmp_path / "contacts.xlsx"
    df.to_excel(xlsx_path, index=False)

    lookup = build_contact_lookup(str(xlsx_path))

    msg = Message(
        date_raw="",
        date_dt=None,
        msg_type="sms",
        direction="in",
        attachments=[],
        body="Hello",
        sender="1234567890",
        recipients="",
        message_id="id1",
        attachment_day=None,
    )

    out_file = tmp_path / "out.html"
    render_thread_html(tmp_path, out_file, [msg], ["1234567890"], "1234567890", lookup)

    html = out_file.read_text()
    # Contact name is shown together with the phone number
    assert '<div class="sender">Alice Smith (1234567890)</div>' in html
    assert "Participants: Alice Smith (1234567890)" in html


@pytest.mark.parametrize(
    "raw,expected",
    [
        ("+12223334444", "2223334444"),
        ("111-222-3333", "1112223333"),
        ("(111) 222-3333", "1112223333"),
        ("+1 111-222-3333", "1112223333"),
        ("1112223333", "1112223333"),
    ],
)
def test_normalize_phone_number_variants(raw, expected):
    assert normalize_phone_number(raw) == expected


def test_contact_lookup_handles_various_phone_formats(tmp_path):
    df = pd.DataFrame(
        [
            {"firstname": "Alice", "lastname": "Smith", "phone_numbers": "(111) 222-3333"},
            {"firstname": "Bob", "lastname": "Jones", "phone_numbers": "+1 222-333-4444"},
        ]
    )
    xlsx_path = tmp_path / "contacts.xlsx"
    df.to_excel(xlsx_path, index=False)

    lookup = build_contact_lookup(str(xlsx_path))

    assert lookup("+12223334444") == "Bob Jones"
    for variant in ["111-222-3333", "(111) 222-3333", "+1 111-222-3333", "1112223333"]:
        assert lookup(variant) == "Alice Smith"
