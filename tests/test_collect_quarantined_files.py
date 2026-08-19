import importlib
import sys
import zipfile
from pathlib import Path


def load_module():
    project_root = Path(__file__).resolve().parents[1]
    sys.path.append(str(project_root))
    return importlib.import_module("synchronoss_parser.collect_quarantined_files")


def test_collect_quarantined_files(tmp_path):
    module = load_module()
    root = tmp_path / "VZMOBILE"
    root.mkdir()
    compiled = root / "Compiled Quarantine Files"

    # Write raw PNG bytes directly to the ``zip_file`` path. The function
    # should recognize the media type and copy it with the correct
    # extension rather than attempting extraction.
    data = b"\x89PNG\r\n\x1a\nrest"
    zip_path = root / "sample.zip_file_1"
    zip_path.write_bytes(data)

    copied, skipped, total = module.collect_quarantined_files(root, compiled)

    assert len(copied) == 1
    assert skipped == []
    assert total == 1
    dest = compiled / "sample.png"
    assert dest.exists()
    assert dest.suffix == ".png"
    assert copied[0] == dest


def test_skips_invalid_archives(tmp_path):
    module = load_module()
    root = tmp_path / "VZMOBILE"
    root.mkdir()
    compiled = root / "Compiled Quarantine Files"

    # Valid media file stored directly under ``zip_file`` naming
    data = b"\xFF\xD8\xFFrest"
    valid_zip = root / "valid.zip_file_1"
    valid_zip.write_bytes(data)

    # Dummy text file with zip_file_ pattern
    invalid_zip = root / "broken.zip_file_2"
    invalid_zip.write_text("not a zip")

    copied, skipped, total = module.collect_quarantined_files(root, compiled)

    dest = compiled / "valid.jpg"
    assert dest.exists()
    assert dest in copied
    assert skipped == [invalid_zip]
    assert total == 2


def test_multiple_independent_files(tmp_path):
    module = load_module()
    root = tmp_path / "VZMOBILE"
    root.mkdir()
    compiled = root / "Compiled Quarantine Files"

    # Two separate media files stored directly as ``zip_file_1``
    (root / "one.zip_file_1").write_bytes(b"\x89PNG\r\n\x1a\nrest")
    (root / "two.zip_file_1").write_bytes(b"\xFF\xD8\xFFrest")

    copied, skipped, total = module.collect_quarantined_files(root, compiled)

    assert skipped == []

    assert len(copied) == 2

    dest_png = compiled / "one.png"
    dest_jpg = compiled / "two.jpg"
    assert dest_png.exists()
    assert dest_jpg.exists()
    assert set(copied) == {dest_png, dest_jpg}
    assert total == 2



def test_skips_unallowed_extension(tmp_path):
    module = load_module()
    root = tmp_path / "VZMOBILE"
    root.mkdir()
    compiled = root / "Compiled Quarantine Files"

    data = b"%PDF-"
    zip_path = root / "doc.zip_file_1"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.writestr("file", data)

    copied, skipped, total = module.collect_quarantined_files(root, compiled)

    assert copied == []
    assert skipped == [zip_path]
    assert list(compiled.iterdir()) == []
    assert total == 1

