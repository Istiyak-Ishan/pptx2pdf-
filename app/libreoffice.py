import logging
import os
import shutil
import winreg
from pathlib import Path

logger = logging.getLogger(__name__)

# Standard candidate paths on Windows
STANDARD_PATHS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]

def is_valid_soffice(executable_path: str | Path) -> bool:
    """Check if the given path points to an existing file."""
    if not executable_path:
        return False
    path = Path(executable_path)
    return path.is_file() and path.name.lower() in ("soffice.exe", "soffice")

def _check_registry() -> Path | None:
    """Look for LibreOffice installation path in the Windows Registry."""
    reg_keys = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\LibreOffice\UNO\InstallPath"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\LibreOffice\UNO\InstallPath"),
    ]

    for root_key, subkey in reg_keys:
        try:
            with winreg.OpenKey(root_key, subkey) as key:
                install_dir, _ = winreg.QueryValueEx(key, "")
                if install_dir:
                    candidate = Path(install_dir) / "soffice.exe"
                    if candidate.is_file():
                        logger.info(f"LibreOffice discovered via Registry: {candidate}")
                        return candidate
        except OSError:
            continue

    return None

def find_libreoffice(custom_path: str | None = None) -> Path | None:
    """
    Search for LibreOffice executable on the system.
    
    Order of preference:
    1. custom_path (if provided and valid)
    2. System PATH (via shutil.which)
    3. Standard installation paths
    4. Windows Registry

    :param custom_path: Optional explicit path to soffice executable provided by user.
    :return: Resolved Path to soffice executable, or None if not found.
    """
    # 1. Check custom user-specified path
    if custom_path:
        if is_valid_soffice(custom_path):
            logger.info(f"Using custom LibreOffice path: {custom_path}")
            return Path(custom_path)
        else:
            logger.warning(f"Custom LibreOffice path provided is invalid or missing: {custom_path}")

    # 2. Check system PATH
    path_executable = shutil.which("soffice") or shutil.which("soffice.exe")
    if path_executable and is_valid_soffice(path_executable):
        logger.info(f"LibreOffice discovered on system PATH: {path_executable}")
        return Path(path_executable)

    # 3. Check standard installation directories
    # Also expand ProgramFiles env vars dynamically
    program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
    program_files_x86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    local_appdata = os.environ.get("LOCALAPPDATA", "")

    dynamic_paths = [
        Path(program_files) / "LibreOffice" / "program" / "soffice.exe",
        Path(program_files_x86) / "LibreOffice" / "program" / "soffice.exe",
    ]
    if local_appdata:
        dynamic_paths.append(Path(local_appdata) / "Programs" / "LibreOffice" / "program" / "soffice.exe")

    all_candidate_paths = [Path(p) for p in STANDARD_PATHS] + dynamic_paths

    for candidate in all_candidate_paths:
        if is_valid_soffice(candidate):
            logger.info(f"LibreOffice discovered at standard path: {candidate}")
            return candidate

    # 4. Check Windows Registry
    registry_path = _check_registry()
    if registry_path:
        return registry_path

    logger.warning("LibreOffice could not be found on this system.")
    return None
