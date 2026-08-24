import pytest
from pathlib import Path
from app.utils import is_valid_pptx, get_unique_pdf_path, open_output_folder

def test_is_valid_pptx_valid(tmp_path: Path):
    pptx_file = tmp_path / "test.pptx"
    pptx_file.write_text("fake pptx content")
    assert is_valid_pptx(pptx_file) is True
    assert is_valid_pptx(str(pptx_file)) is True

def test_is_valid_pptx_invalid_extension(tmp_path: Path):
    txt_file = tmp_path / "test.txt"
    txt_file.write_text("text content")
    assert is_valid_pptx(txt_file) is False

def test_is_valid_pptx_nonexistent():
    assert is_valid_pptx("nonexistent_file_12345.pptx") is False
    assert is_valid_pptx("") is False
    assert is_valid_pptx(None) is False

def test_get_unique_pdf_path_no_collision(tmp_path: Path):
    target = get_unique_pdf_path(tmp_path, "presentation")
    assert target == tmp_path / "presentation.pdf"

def test_get_unique_pdf_path_with_collision(tmp_path: Path):
    (tmp_path / "presentation.pdf").touch()
    target1 = get_unique_pdf_path(tmp_path, "presentation")
    assert target1 == tmp_path / "presentation (1).pdf"

    target1.touch()
    target2 = get_unique_pdf_path(tmp_path, "presentation")
    assert target2 == tmp_path / "presentation (2).pdf"

def test_open_output_folder_nonexistent():
    success, err = open_output_folder("C:/nonexistent_folder_path_xyz_99")
    assert success is False
    assert "does not exist" in err.lower() or "not a directory" in err.lower()

def test_get_log_filepath():
    from app.config import get_log_filepath
    log_path = get_log_filepath("test_custom.log")
    assert log_path.name == "test_custom.log"
    assert log_path.parent.name == "logs"

