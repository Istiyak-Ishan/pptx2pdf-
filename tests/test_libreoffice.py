from unittest.mock import patch

from app.libreoffice import find_libreoffice, is_valid_soffice


def test_is_valid_soffice_invalid():
    assert is_valid_soffice(None) is False
    assert is_valid_soffice("") is False
    assert is_valid_soffice("non_existent_file.exe") is False

def test_is_valid_soffice_valid(tmp_path):
    soffice_file = tmp_path / "soffice.exe"
    soffice_file.write_text("dummy binary")
    assert is_valid_soffice(soffice_file) is True

def test_find_libreoffice_custom_path_valid(tmp_path):
    custom = tmp_path / "soffice.exe"
    custom.write_text("dummy")
    result = find_libreoffice(custom_path=str(custom))
    assert result == custom

def test_find_libreoffice_path_fallback(tmp_path):
    path_executable = tmp_path / "soffice.exe"
    path_executable.write_text("dummy")

    with patch("shutil.which", return_value=str(path_executable)):
        result = find_libreoffice()
        assert result == path_executable

def test_find_libreoffice_not_found():
    with patch("shutil.which", return_value=None), \
         patch("app.libreoffice.is_valid_soffice", return_value=False), \
         patch("app.libreoffice._check_registry", return_value=None):
        result = find_libreoffice()
        assert result is None
