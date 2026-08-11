import logging
import os
import sys
from pathlib import Path

logger = logging.getLogger(__name__)

def is_valid_pptx(path_str: str | Path) -> bool:
    """Check if file exists and has a .pptx extension."""
    if not path_str:
        return False
    path = Path(path_str)
    return path.is_file() and path.suffix.lower() == ".pptx"

def get_unique_pdf_path(output_folder: Path, stem: str) -> Path:
    """
    Generate a non-colliding PDF output path in the given folder.
    e.g. presentation.pdf -> presentation (1).pdf -> presentation (2).pdf
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
    Open output folder in Windows File Explorer.
    Returns (success, error_message).
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
