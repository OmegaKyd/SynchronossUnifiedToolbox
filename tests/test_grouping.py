import sys
from pathlib import Path

import pytest

sys.path.append(str(Path(__file__).resolve().parents[1]))
from synchronoss_parser.render_transcripts import Message, group_messages_by_chat


def make_message(direction, sender="", recipients="", msg_type="sms"):
    return Message(
        date_raw="",
        date_dt=None,
        msg_type=msg_type,
        direction=direction,
        attachments=[],
        body="",
        sender=sender,
        recipients=recipients,
        message_id="id",
        attachment_day=None,
    )


def test_one_on_one_grouping():
    msgs = [
        make_message("in", sender="111", recipients=""),
        make_message("out", sender="", recipients="111"),
    ]
    groups = group_messages_by_chat(msgs, target="222")
    assert len(groups) == 1
    key = next(iter(groups))
    assert key == ("111", "222")
    grouped_msgs = groups[key]
    assert grouped_msgs[0].recipients == "222"
    assert grouped_msgs[1].sender == "222"


def test_group_chat_grouping():
    msgs = [
        make_message("out", sender="", recipients="333;444"),
        make_message("in", sender="333", recipients="444"),
    ]
    groups = group_messages_by_chat(msgs, target="222")
    assert len(groups) == 1
    key = next(iter(groups))
    assert key == ("222", "333", "444")
    grouped_msgs = groups[key]
    assert grouped_msgs[0].sender == "222"
    assert grouped_msgs[1].recipients == "444"
