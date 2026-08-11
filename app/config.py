import logging
import sys
from pathlib import Path

APP_NAME = "PPTX2PDF"
APP_TITLE = "PPTX2PDF — PowerPoint to PDF Converter"
APP_VERSION = "1.0.0"

def get_base_dir() -> Path:
    """Get the base directory of the application (handles PyInstaller bundle)."""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent.parent

def get_asset_path(filename: str) -> Path:
    """Get the absolute path to an asset file."""
    return get_base_dir() / "assets" / filename

def get_logs_dir() -> Path:
    """Get or create the directory where logs should be stored."""
    if getattr(sys, 'frozen', False):
        # Executable mode: store logs relative to executable location or user AppData
        exe_dir = Path(sys.executable).parent
        logs_dir = exe_dir / "logs"
    else:
        logs_dir = get_base_dir() / "logs"
    
    logs_dir.mkdir(parents=True, exist_ok=True)
    return logs_dir

def setup_logging(log_filename: str = "pptx2pdf.log") -> Path:
    """Set up application logging."""
    logs_dir = get_logs_dir()
    log_file = logs_dir / log_filename

    # Configure root logger
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logger = logging.getLogger(__name__)
    logger.info(f"Logging initialized. Application version: {APP_VERSION}")
    return log_file
