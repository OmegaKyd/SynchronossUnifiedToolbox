import importlib
import sys
from pathlib import Path


def load_module():
    project_root = Path(__file__).resolve().parents[1]
    sys.path.append(str(project_root))
    return importlib.import_module("synchronoss_parser.collect_quarantined_files")


def test_detect_mp3(tmp_path):
    module = load_module()
    data = b"ID3" + b"\x00" * 13
    path = tmp_path / "audio"
    path.write_bytes(data)
    assert module.detect_extension(path) == ".mp3"


def test_detect_wav(tmp_path):
    module = load_module()
    data = b"RIFF" + b"\x00" * 4 + b"WAVE" + b"\x00" * 4
    path = tmp_path / "audio"
    path.write_bytes(data)
    assert module.detect_extension(path) == ".wav"


def test_detect_mov(tmp_path):
    module = load_module()
    data = b"\x00\x00\x00\x14ftypqt  " + b"\x00" * 4
    path = tmp_path / "video"
    path.write_bytes(data)
    assert module.detect_extension(path) == ".mov"
