import sys
from pathlib import Path

from PIL import Image

sys.path.append(str(Path(__file__).resolve().parents[1]))
from synchronoss_parser.attachment_log import create_thumbnail

def test_create_thumbnail_non_image(tmp_path):
    src = tmp_path / "file.txt"
    src.write_text("not an image")
    dest = tmp_path / "thumb.png"
    assert create_thumbnail(src, dest) is False
    assert not dest.exists()

def test_create_thumbnail_image(tmp_path):
    src = tmp_path / "image.png"
    Image.new("RGB", (10, 10), color="red").save(src)
    dest = tmp_path / "thumb.png"
    assert create_thumbnail(src, dest) is True
    assert dest.exists()
