import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1]))
from synchronoss_parser.render_transcripts import Message, render_thread_html


def test_empty_sms_renders_placeholder(tmp_path):
    msg = Message(
        date_raw="",
        date_dt=None,
        msg_type="sms",
        direction="in",
        attachments=[],
        body="",
        sender="123",
        recipients="",
        message_id="id1",
        attachment_day=None,
    )

    out_file = tmp_path / "out.html"
    total, with_attach = render_thread_html(tmp_path, out_file, [msg], ["123"], "123")

    assert total == 1
    assert with_attach == 0

    html = out_file.read_text()
    placeholder = "NO DATA IN CSV FOR THIS SMS MESSAGE - LOG ONLY"
    assert f'<div class="body-text missing">{placeholder}</div>' in html
