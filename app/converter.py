import logging
import shutil
import subprocess
import tempfile
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

from app.libreoffice import find_libreoffice, is_valid_soffice
from app.utils import get_unique_pdf_path, is_valid_pptx

logger = logging.getLogger(__name__)

@dataclass
class ConversionResult:
    input_path: Path
    output_path: Path | None = None
    success: bool = False
    error_message: str = ""
    duration_seconds: float = 0.0

@dataclass
class BatchConversionResult:
    total: int = 0
    successful_count: int = 0
    failed_count: int = 0
    results: list[ConversionResult] = field(default_factory=list)
    total_duration: float = 0.0

class PPTXConverter:
    """Core PPTX to PDF conversion engine utilizing LibreOffice in headless mode."""

    def __init__(self, libreoffice_path: str | Path | None = None) -> None:
        if libreoffice_path:
            self.libreoffice_path = Path(libreoffice_path) if is_valid_soffice(libreoffice_path) else None
        else:
            self.libreoffice_path = find_libreoffice()

    def set_libreoffice_path(self, path: str | Path) -> bool:
        """Update LibreOffice executable path."""
        if is_valid_soffice(path):
            self.libreoffice_path = Path(path)
            return True
        return False

    def find_and_set_libreoffice(self) -> bool:
        """Re-scan system for LibreOffice executable."""
        discovered = find_libreoffice()
        if discovered:
            self.libreoffice_path = discovered
            return True
        return False

    def is_ready(self) -> bool:
        """Check if converter has a valid LibreOffice executable."""
        return self.libreoffice_path is not None and is_valid_soffice(self.libreoffice_path)

    def convert_file(
        self,
        input_file: str | Path,
        output_folder: str | Path,
        overwrite: bool = True,
        timeout_seconds: int = 300
    ) -> ConversionResult:
        """
        Convert a single PPTX file to PDF.

        :param input_file: Path to source .pptx file.
        :param output_folder: Destination folder for converted .pdf.
        :param overwrite: Whether to overwrite existing destination files.
        :param timeout_seconds: Maximum time allowed for subprocess conversion.
        :return: ConversionResult object containing status and metadata.
        """
        start_time = time.time()
        input_path = Path(input_file).resolve()
        out_folder_path = Path(output_folder).resolve()

        result = ConversionResult(input_path=input_path)

        # 1. Check LibreOffice installation
        if not self.is_ready():
            result.error_message = "LibreOffice executable was not found."
            logger.error(result.error_message)
            return result

        # 2. Validate input file
        if not input_path.exists():
            result.error_message = f"Input file does not exist: {input_path}"
            logger.error(result.error_message)
            return result

        if not is_valid_pptx(input_path):
            result.error_message = f"Invalid file format (must be .pptx): {input_path.name}"
            logger.error(result.error_message)
            return result

        # 3. Validate output directory
        try:
            out_folder_path.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            result.error_message = f"Failed to create output directory {out_folder_path}: {e}"
            logger.exception(result.error_message)
            return result

        # Determine target PDF path in final destination
        stem = input_path.stem
        if overwrite:
            target_pdf_path = out_folder_path / f"{stem}.pdf"
        else:
            target_pdf_path = get_unique_pdf_path(out_folder_path, stem)

        # 4. Use a isolated temp directory for LibreOffice --outdir to avoid file conflicts
        with tempfile.TemporaryDirectory(prefix="pptx2pdf_") as temp_dir:
            temp_dir_path = Path(temp_dir)
            cmd = [
                str(self.libreoffice_path),
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(temp_dir_path),
                str(input_path)
            ]

            logger.info(f"Starting conversion: '{input_path.name}' -> PDF")
            logger.debug(f"Executing command: {cmd}")

            try:
                completed = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    timeout=timeout_seconds,
                    check=False
                )

                if completed.returncode != 0:
                    err = completed.stderr.strip() or completed.stdout.strip() or f"Process exited with code {completed.returncode}"
                    result.error_message = f"LibreOffice conversion failed: {err}"
                    logger.error(f"Conversion error for {input_path.name}: {result.error_message}")
                    return result

                # Check for output PDF in temp directory
                temp_pdf = temp_dir_path / f"{stem}.pdf"
                if not temp_pdf.exists():
                    # Check if any .pdf file was generated in temp_dir
                    pdfs = list(temp_dir_path.glob("*.pdf"))
                    if pdfs:
                        temp_pdf = pdfs[0]
                    else:
                        result.error_message = "Conversion command completed but output PDF was not created."
                        logger.error(result.error_message)
                        return result

                # Copy/Move temp PDF to target path in output_folder
                shutil.move(str(temp_pdf), str(target_pdf_path))
                result.success = True
                result.output_path = target_pdf_path
                result.duration_seconds = round(time.time() - start_time, 2)
                logger.info(f"Successfully converted '{input_path.name}' in {result.duration_seconds}s -> {target_pdf_path}")
                return result

            except subprocess.TimeoutExpired:
                result.error_message = f"Conversion timed out after {timeout_seconds} seconds."
                logger.error(result.error_message)
                return result
            except Exception as e:
                result.error_message = f"Unexpected error during conversion: {e!s}"
                logger.exception(result.error_message)
                return result

    def convert_batch(
        self,
        input_files: list[str | Path],
        output_folder: str | Path,
        overwrite: bool = True,
        timeout_seconds: int = 300,
        progress_callback: Callable[[int, int, str, str], None] | None = None
    ) -> BatchConversionResult:
        """
        Convert multiple PPTX files to PDF.
        
        progress_callback signature:
            callback(current_index, total_count, filename, status_message)
        """
        batch_start_time = time.time()
        paths = [Path(f) for f in input_files]
        batch_result = BatchConversionResult(total=len(paths))

        for idx, input_path in enumerate(paths, start=1):
            filename = input_path.name
            if progress_callback:
                progress_callback(idx, len(paths), filename, f"Converting {filename} ({idx}/{len(paths)})...")

            res = self.convert_file(
                input_path,
                output_folder,
                overwrite=overwrite,
                timeout_seconds=timeout_seconds
            )
            batch_result.results.append(res)

            if res.success:
                batch_result.successful_count += 1
            else:
                batch_result.failed_count += 1

        batch_result.total_duration = round(time.time() - batch_start_time, 2)
        logger.info(
            f"Batch conversion completed. Successful: {batch_result.successful_count}/{batch_result.total}, "
            f"Failed: {batch_result.failed_count}, Duration: {batch_result.total_duration}s"
        )
        return batch_result
