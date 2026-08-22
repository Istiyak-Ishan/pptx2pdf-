import argparse
import logging
import sys
from pathlib import Path

from app.config import APP_TITLE, APP_VERSION, setup_logging
from app.converter import PPTXConverter
from app.utils import open_output_folder


def run_cli(args: argparse.Namespace) -> int:
    """Run conversion in CLI mode."""
    setup_logging("pptx2pdf_cli.log")
    logger = logging.getLogger(__name__)

    input_paths = [Path(p) for p in args.inputs]
    output_dir = Path(args.output) if args.output else input_paths[0].parent

    logger.info(f"CLI mode invoked for {len(input_paths)} file(s). Output: {output_dir}")
    print(f"[{APP_TITLE}] Starting conversion...")

    converter = PPTXConverter(libreoffice_path=args.libreoffice)
    if not converter.is_ready():
        print("Error: LibreOffice was not found. Please install LibreOffice or specify --libreoffice path.")
        return 1

    result = converter.convert_batch(
        input_files=input_paths,
        output_folder=output_dir,
        overwrite=not args.no_overwrite,
        progress_callback=lambda idx, total, fname, msg: print(f"[{idx}/{total}] {msg}")
    )

    print("-" * 50)
    print(f"Conversion complete! Successful: {result.successful_count}/{result.total}, Failed: {result.failed_count}")
    print(f"Total time: {result.total_duration} seconds")

    if result.failed_count > 0:
        print("\nFailures:")
        for res in result.results:
            if not res.success:
                print(f"  • {res.input_path.name}: {res.error_message}")
        return 1

    if args.open:
        open_output_folder(output_dir)

    return 0

def run_gui() -> None:
    """Launch the CustomTkinter desktop GUI."""
    setup_logging("pptx2pdf.log")
    logger = logging.getLogger(__name__)
    logger.info(f"Launching {APP_TITLE} GUI...")

    from app.gui import PPTX2PDFApp
    app = PPTX2PDFApp()
    app.mainloop()

def main() -> None:
    """Main application entry point."""
    parser = argparse.ArgumentParser(
        prog="pptx2pdf",
        description=f"{APP_TITLE} — Convert PowerPoint .pptx files to PDF using LibreOffice."
    )
    parser.add_argument("inputs", nargs="*", help="One or more .pptx files to convert.")
    parser.add_argument("-o", "--output", help="Output folder directory.")
    parser.add_argument("--libreoffice", help="Custom path to LibreOffice soffice.exe.")
    parser.add_argument("--no-overwrite", action="store_true", help="Do not overwrite existing PDF files.")
    parser.add_argument("--open", action="store_true", help="Open output folder after conversion.")
    parser.add_argument("--version", action="version", version=f"{APP_TITLE} v{APP_VERSION}")

    args = parser.parse_args()

    # If input files were provided on CLI, run CLI mode. Otherwise launch GUI.
    if args.inputs:
        sys.exit(run_cli(args))
    else:
        run_gui()

if __name__ == "__main__":
    main()
