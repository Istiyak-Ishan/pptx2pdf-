import logging
import os
import sys
from pathlib import Path

logger = logging.getLogger(__name__)

def is_valid_pptx(path_str: str | Path | None) -> bool:
    """
    Check if the given path string or Path object exists and has a .pptx extension.

    :param path_str: Path object or string representation of file path.
    :return: True if valid existing .pptx file, False otherwise.
    """
    if not path_str:
        return False
    path = Path(path_str)
    return path.is_file() and path.suffix.lower() == ".pptx"

def get_unique_pdf_path(output_folder: Path, stem: str) -> Path:
    """
    Generate a non-colliding PDF output path in the given folder.
    e.g. presentation.pdf -> presentation (1).pdf -> presentation (2).pdf

    :param output_folder: Destination folder Path object.
    :param stem: Base filename stem without extension.
    :return: Resolved non-colliding output Path object.
    """
    target = output_folder / f"{stem}.pdf"
    if not target.exists():
        return target

    counter = 1
    while True:
        candidate = output_folder / f"{stem} ({counter}).pdf"
        if not candidate.exists():
            return candidate
        counter += 1

def open_output_folder(folder_path: str | Path) -> tuple[bool, str]:
    """
    Open output folder in system file manager (Windows File Explorer or OS equivalent).

    :param folder_path: Target directory path to open.
    :return: Tuple of (success_boolean, error_message_string).
    """
    try:
        path = Path(folder_path).resolve()
        if not path.exists():
            return False, f"Folder does not exist: {path}"
        if not path.is_dir():
            return False, f"Path is not a directory: {path}"

        if sys.platform == "win32":
            os.startfile(str(path))
            return True, ""
        else:
            # Fallback for non-windows OS in development environments
            import subprocess
            if sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
            return True, ""
    except Exception as e:
        logger.exception(f"Failed to open folder {folder_path}")
        return False, str(e)


def format_file_size(size_bytes: float) -> str:
    """
    Format a byte count into a human-readable string (e.g. B, KB, MB, GB).

    :param size_bytes: Size in bytes.
    :return: Formatted string representation of size.
    """
    if size_bytes < 0:
        return "0 B"
    size = float(size_bytes)
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size < 1024.0 or unit == "TB":
            if unit == "B":
                return f"{int(size)} B"
            return f"{size:.1f} {unit}"
        size /= 1024.0
    return f"{size:.1f} PB"

