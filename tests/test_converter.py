from pathlib import Path
from unittest.mock import MagicMock, patch

# pyrefly: ignore [missing-import]
import pytest

from app.converter import PPTXConverter


@pytest.fixture
def mock_converter(tmp_path):
    soffice_path = tmp_path / "soffice.exe"
    soffice_path.write_text("dummy binary")
    converter = PPTXConverter(libreoffice_path=soffice_path)
    return converter, soffice_path

def test_converter_not_ready():
    converter = PPTXConverter(libreoffice_path="invalid_path.exe")
    assert converter.is_ready() is False

def test_invalid_input_file(mock_converter, tmp_path):
    converter, _ = mock_converter
    out_dir = tmp_path / "out"

    # Non-existent file
    res = converter.convert_file(tmp_path / "missing.pptx", out_dir)
    assert res.success is False
    assert "does not exist" in res.error_message

    # Invalid extension
    bad_file = tmp_path / "test.txt"
    bad_file.write_text("hello")
    res2 = converter.convert_file(bad_file, out_dir)
    assert res2.success is False
    assert "Invalid file format" in res2.error_message

def test_single_file_conversion_success(mock_converter, tmp_path):
    converter, _ = mock_converter
    pptx_file = tmp_path / "presentation.pptx"
    pptx_file.write_bytes(b"dummy pptx content")
    out_dir = tmp_path / "out"

    # Mock subprocess.run to simulate LibreOffice creating presentation.pdf in temp_dir
    def mock_subprocess_run(cmd, capture_output, text, timeout, check):
        # find outdir from cmd list
        outdir_idx = cmd.index("--outdir") + 1
        temp_dir = Path(cmd[outdir_idx])
        # create mock pdf
        pdf = temp_dir / "presentation.pdf"
        pdf.write_bytes(b"%PDF-1.4 dummy pdf content")
        mock_proc = MagicMock()
        mock_proc.returncode = 0
        mock_proc.stdout = ""
        mock_proc.stderr = ""
        return mock_proc

    with patch("subprocess.run", side_effect=mock_subprocess_run):
        res = converter.convert_file(pptx_file, out_dir)
        assert res.success is True
        assert res.output_path == out_dir / "presentation.pdf"
        assert (out_dir / "presentation.pdf").exists()

def test_filename_with_spaces_and_bangla(mock_converter, tmp_path):
    converter, _ = mock_converter
    bangla_file = tmp_path / "DS Lab presentation বাংলা.pptx"
    bangla_file.write_bytes(b"dummy pptx")
    out_dir = tmp_path / "out"

    def mock_subprocess_run(cmd, capture_output, text, timeout, check):
        outdir_idx = cmd.index("--outdir") + 1
        temp_dir = Path(cmd[outdir_idx])
        pdf = temp_dir / f"{bangla_file.stem}.pdf"
        pdf.write_bytes(b"%PDF-1.4 dummy pdf content")
        mock_proc = MagicMock()
        mock_proc.returncode = 0
        return mock_proc

    with patch("subprocess.run", side_effect=mock_subprocess_run):
        res = converter.convert_file(bangla_file, out_dir)
        assert res.success is True
        assert res.output_path.name == "DS Lab presentation বাংলা.pdf"

def test_existing_pdf_no_overwrite(mock_converter, tmp_path):
    converter, _ = mock_converter
    pptx_file = tmp_path / "doc.pptx"
    pptx_file.write_bytes(b"dummy pptx")
    out_dir = tmp_path / "out"
    out_dir.mkdir(parents=True)

    # Existing doc.pdf
    existing_pdf = out_dir / "doc.pdf"
    existing_pdf.write_bytes(b"existing content")

    def mock_subprocess_run(cmd, capture_output, text, timeout, check):
        outdir_idx = cmd.index("--outdir") + 1
        temp_dir = Path(cmd[outdir_idx])
        pdf = temp_dir / "doc.pdf"
        pdf.write_bytes(b"%PDF-1.4 new pdf content")
        mock_proc = MagicMock()
        mock_proc.returncode = 0
        return mock_proc

    with patch("subprocess.run", side_effect=mock_subprocess_run):
        # Convert with overwrite=False
        res = converter.convert_file(pptx_file, out_dir, overwrite=False)
        assert res.success is True
        assert res.output_path == out_dir / "doc (1).pdf"
        assert existing_pdf.read_bytes() == b"existing content"

def test_batch_conversion(mock_converter, tmp_path):
    converter, _ = mock_converter
    f1 = tmp_path / "pres1.pptx"
    f2 = tmp_path / "pres2.pptx"
    f1.write_bytes(b"p1")
    f2.write_bytes(b"p2")
    out_dir = tmp_path / "out"

    def mock_subprocess_run(cmd, capture_output, text, timeout, check):
        outdir_idx = cmd.index("--outdir") + 1
        temp_dir = Path(cmd[outdir_idx])
        inp_stem = Path(cmd[-1]).stem
        pdf = temp_dir / f"{inp_stem}.pdf"
        pdf.write_bytes(b"%PDF-1.4")
        mock_proc = MagicMock()
        mock_proc.returncode = 0
        return mock_proc

    progress_calls = []
    def progress_cb(idx, total, fname, msg):
        progress_calls.append((idx, total, fname))

    with patch("subprocess.run", side_effect=mock_subprocess_run):
        batch_res = converter.convert_batch(
            input_files=[f1, f2],
            output_folder=out_dir,
            progress_callback=progress_cb
        )

        assert batch_res.total == 2
        assert batch_res.successful_count == 2
        assert batch_res.failed_count == 0
        assert len(progress_calls) == 2

def test_partial_batch_failure_recovery(mock_converter, tmp_path):
    converter, _ = mock_converter
    f_good = tmp_path / "good.pptx"
    f_bad = tmp_path / "bad.pptx"
    f_good.write_bytes(b"good")
    f_bad.write_bytes(b"bad")
    out_dir = tmp_path / "out"

    def mock_subprocess_run(cmd, capture_output, text, timeout, check):
        inp_stem = Path(cmd[-1]).stem
        outdir_idx = cmd.index("--outdir") + 1
        temp_dir = Path(cmd[outdir_idx])
        mock_proc = MagicMock()

        if inp_stem == "bad":
            mock_proc.returncode = 1
            mock_proc.stderr = "Corrupted presentation"
        else:
            pdf = temp_dir / f"{inp_stem}.pdf"
            pdf.write_bytes(b"%PDF-1.4")
            mock_proc.returncode = 0

        return mock_proc

    with patch("subprocess.run", side_effect=mock_subprocess_run):
        batch_res = converter.convert_batch(
            input_files=[f_good, f_bad],
            output_folder=out_dir
        )

        assert batch_res.total == 2
        assert batch_res.successful_count == 1
        assert batch_res.failed_count == 1
        assert batch_res.results[1].success is False
        assert "Corrupted presentation" in batch_res.results[1].error_message

def test_batch_conversion_with_custom_timeout(mock_converter, tmp_path):
    converter, _ = mock_converter
    f1 = tmp_path / "test.pptx"
    f1.write_bytes(b"content")
    out_dir = tmp_path / "out"

    captured_timeout = []
    def mock_subprocess_run(cmd, capture_output, text, timeout, check):
        captured_timeout.append(timeout)
        outdir_idx = cmd.index("--outdir") + 1
        temp_dir = Path(cmd[outdir_idx])
        pdf = temp_dir / "test.pdf"
        pdf.write_bytes(b"%PDF-1.4")
        mock_proc = MagicMock()
        mock_proc.returncode = 0
        return mock_proc

    with patch("subprocess.run", side_effect=mock_subprocess_run):
        batch_res = converter.convert_batch(
            input_files=[f1],
            output_folder=out_dir,
            timeout_seconds=45
        )
        assert batch_res.successful_count == 1
        assert captured_timeout == [45]

