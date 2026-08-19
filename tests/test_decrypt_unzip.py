from __future__ import annotations

import importlib
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest


def load_module():
    project_root = Path(__file__).resolve().parents[1]
    sys.path.append(str(project_root))
    return importlib.import_module("synchronoss_parser.decrypt_unzip")


PASSWORD = "s3cr3t"


def test_decrypt_and_unzip(tmp_path: Path) -> None:
    """Decrypt nested archives and ensure all files are extracted."""
    if shutil.which("gpg") is None:
        pytest.skip("gpg not installed")

    decrypt_unzip = load_module()

    # Create a plain zipped file
    plain_zip = tmp_path / "plain.zip"
    with zipfile.ZipFile(plain_zip, "w") as zf:
        zf.writestr("plain.txt", "hello")

    # Create a zipped file and encrypt it with gpg
    secret_zip = tmp_path / "secret.zip"
    with zipfile.ZipFile(secret_zip, "w") as zf:
        zf.writestr("secret.txt", "topsecret")

    secret_gpg = tmp_path / "secret.zip.gpg"
    subprocess.run(
        [
            "gpg",
            "--batch",
            "--yes",
            "--passphrase",
            PASSWORD,
            "--symmetric",
            "-o",
            str(secret_gpg),
            str(secret_zip),
        ],
        check=True,
    )

    # Bundle both files into a parent archive
    archive_zip = tmp_path / "archive.zip"
    with zipfile.ZipFile(archive_zip, "w") as zf:
        zf.write(plain_zip, plain_zip.name)
        zf.write(secret_gpg, secret_gpg.name)

    # Decrypt and unzip the archive
    files = decrypt_unzip.decrypt_and_unzip(archive_zip, PASSWORD)

    archive = tmp_path / "archive"
    extracted_plain = archive / "plain" / "plain.txt"
    extracted_secret = archive / "secret" / "secret.txt"

    assert set(files) == {extracted_plain, extracted_secret}
    assert extracted_plain.read_text() == "hello"
    assert extracted_secret.read_text() == "topsecret"
    assert not list(archive.glob("**/*.zip"))
    assert not list(archive.glob("**/*.gpg"))
    assert archive_zip.exists()


def test_decrypt_and_unzip_with_strings(tmp_path: Path) -> None:
    """Accept string paths and a custom output directory."""
    if shutil.which("gpg") is None:
        pytest.skip("gpg not installed")

    decrypt_unzip = load_module()

    # Create a plain zipped file
    plain_zip = tmp_path / "plain.zip"
    with zipfile.ZipFile(plain_zip, "w") as zf:
        zf.writestr("plain.txt", "hello")

    # Create a zipped file and encrypt it with gpg
    secret_zip = tmp_path / "secret.zip"
    with zipfile.ZipFile(secret_zip, "w") as zf:
        zf.writestr("secret.txt", "topsecret")

    secret_gpg = tmp_path / "secret.zip.gpg"
    subprocess.run(
        [
            "gpg",
            "--batch",
            "--yes",
            "--passphrase",
            PASSWORD,
            "--symmetric",
            "-o",
            str(secret_gpg),
            str(secret_zip),
        ],
        check=True,
    )

    # Bundle both files into a parent archive
    archive_zip = tmp_path / "archive.zip"
    with zipfile.ZipFile(archive_zip, "w") as zf:
        zf.write(plain_zip, plain_zip.name)
        zf.write(secret_gpg, secret_gpg.name)

    output_dir = tmp_path / "custom_output"

    # Decrypt and unzip the archive using string inputs
    files = decrypt_unzip.decrypt_and_unzip(str(archive_zip), PASSWORD, str(output_dir))

    extracted_plain = output_dir / "plain" / "plain.txt"
    extracted_secret = output_dir / "secret" / "secret.txt"

    assert set(files) == {extracted_plain, extracted_secret}
    assert extracted_plain.read_text() == "hello"
    assert extracted_secret.read_text() == "topsecret"
    assert not list(output_dir.glob("**/*.zip"))
    assert not list(output_dir.glob("**/*.gpg"))
    assert archive_zip.exists()


def test_decrypt_and_unzip_no_cleanup(tmp_path: Path) -> None:
    """Keep original archives when cleanup is disabled."""
    if shutil.which("gpg") is None:
        pytest.skip("gpg not installed")

    decrypt_unzip = load_module()

    # Create a plain zipped file
    plain_zip = tmp_path / "plain.zip"
    with zipfile.ZipFile(plain_zip, "w") as zf:
        zf.writestr("plain.txt", "hello")

    # Create a zipped file and encrypt it with gpg
    secret_zip = tmp_path / "secret.zip"
    with zipfile.ZipFile(secret_zip, "w") as zf:
        zf.writestr("secret.txt", "topsecret")

    secret_gpg = tmp_path / "secret.zip.gpg"
    subprocess.run(
        [
            "gpg",
            "--batch",
            "--yes",
            "--passphrase",
            PASSWORD,
            "--symmetric",
            "-o",
            str(secret_gpg),
            str(secret_zip),
        ],
        check=True,
    )

    # Bundle both files into a parent archive
    archive_zip = tmp_path / "archive.zip"
    with zipfile.ZipFile(archive_zip, "w") as zf:
        zf.write(plain_zip, plain_zip.name)
        zf.write(secret_gpg, secret_gpg.name)

    # Decrypt and unzip the archive without cleaning up
    files = decrypt_unzip.decrypt_and_unzip(archive_zip, PASSWORD, cleanup=False)

    archive = tmp_path / "archive"
    extracted_plain = archive / "plain" / "plain.txt"
    extracted_secret = archive / "secret" / "secret.txt"

    assert set(files) == {extracted_plain, extracted_secret}
    assert extracted_plain.read_text() == "hello"
    assert extracted_secret.read_text() == "topsecret"
    assert list(archive.glob("**/*.zip"))
    assert list(archive.glob("**/*.gpg"))
    assert archive_zip.exists()


def test_decrypt_and_unzip_plain_gpg(tmp_path: Path) -> None:
    """Decrypt standalone ``.gpg`` files inside the archive."""
    if shutil.which("gpg") is None:
        pytest.skip("gpg not installed")

    decrypt_unzip = load_module()

    secret_txt = tmp_path / "secret.txt"
    secret_txt.write_text("topsecret")

    secret_gpg = tmp_path / "secret.txt.gpg"
    subprocess.run(
        [
            "gpg",
            "--batch",
            "--yes",
            "--passphrase",
            PASSWORD,
            "--symmetric",
            "-o",
            str(secret_gpg),
            str(secret_txt),
        ],
        check=True,
    )

    archive_zip = tmp_path / "archive.zip"
    with zipfile.ZipFile(archive_zip, "w") as zf:
        zf.write(secret_gpg, secret_gpg.name)

    files = decrypt_unzip.decrypt_and_unzip(archive_zip, PASSWORD)

    archive = tmp_path / "archive"
    extracted_secret = archive / "secret.txt"

    assert files == [extracted_secret]
    assert extracted_secret.read_text() == "topsecret"
    assert not list(archive.glob("**/*.gpg"))
    assert archive_zip.exists()
