import importlib
import sys
from pathlib import Path
import os


def load_module():
    project_root = Path(__file__).resolve().parents[1]
    sys.path.append(str(project_root))
    return importlib.import_module("synchronoss_parser.collect_quarantined_files")


def test_long_paths(tmp_path, monkeypatch):
    module = load_module()
    root = tmp_path / "VZMOBILE"
    root.mkdir()

    compiled = root / ("Compiled Quarantine Files" + "x" * 100)

    deep = root
    for i in range(10):
        deep = deep / ("dir" + str(i) + "a" * 20)
    deep.mkdir(parents=True)

    # Store raw JPEG bytes directly; the function should handle the long path
    # and rename the file based on its signature without needing extraction.
    zip_path = deep / "sample.zip_file_1"
    data = b"\xFF\xD8\xFFrest"
    zip_path.write_bytes(data)

    # enforce a small path limit to trigger shortening logic
    monkeypatch.setattr(os, "pathconf", lambda *args, **kwargs: 200)

    copied, skipped, total = module.collect_quarantined_files(root, compiled)
    assert skipped == []
    assert len(copied) == 1
    assert copied[0].exists()
    assert len(str(copied[0])) <= 200
    assert total == 1
