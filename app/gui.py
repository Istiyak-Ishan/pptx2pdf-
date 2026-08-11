import logging
import os
import threading
from pathlib import Path
from tkinter import filedialog, messagebox

# pyrefly: ignore [missing-import]
import customtkinter as ctk

from app.config import APP_TITLE, APP_VERSION, get_asset_path
from app.converter import BatchConversionResult, PPTXConverter
from app.utils import is_valid_pptx, open_output_folder

logger = logging.getLogger(__name__)

# CustomTkinter aesthetic configuration
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Theme colors
ACCENT_ORANGE = "#D83B01"
ACCENT_HOVER = "#B02D00"
ACCENT_BLUE = "#0288D1"

class PPTX2PDFApp(ctk.CTk):
    """Main Application Window for PPTX2PDF with modern aesthetic UI."""

    def __init__(self):
        super().__init__()

        self.title(f"{APP_TITLE} v{APP_VERSION}")
        self.geometry("780x670")
        self.minsize(700, 600)

        # Set Window Icon
        icon_path = get_asset_path("icon.ico")
        if icon_path.exists():
            try:
                self.iconbitmap(str(icon_path))
            except (AttributeError, OSError, Exception) as e:  # noqa: BLE001
                logger.warning(f"Could not load iconbitmap: {e}")

        # State Variables
        self.selected_files: list[Path] = []
        self.is_converting: bool = False

        # Initialize Converter
        self.converter = PPTXConverter()

        # Build UI Components
        self._create_header()
        self._create_file_list_section()
        self._create_output_section()
        self._create_settings_section()
        self._create_action_progress_section()
        self._create_footer()

        # Initial LibreOffice Check
        self.after(200, self._check_libreoffice_on_startup)

    def _create_header(self):
        """Header with app title, logo badge, and subtitle banner."""
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", padx=24, pady=(18, 8))

        top_row = ctk.CTkFrame(header_frame, fg_color="transparent")
        top_row.pack(fill="x")

        title_label = ctk.CTkLabel(
            top_row,
            text="PPTX2PDF",
            font=ctk.CTkFont(size=26, weight="bold")
        )
        title_label.pack(side="left")

        # Accent Pill Badge
        badge = ctk.CTkFrame(top_row, fg_color=ACCENT_ORANGE, corner_radius=12)
        badge.pack(side="left", padx=(10, 0))
        badge_lbl = ctk.CTkLabel(
            badge,
            text="PPTX ➔ PDF",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="white"
        )
        badge_lbl.pack(padx=10, pady=2)

        ver_lbl = ctk.CTkLabel(
            top_row,
            text=f"v{APP_VERSION}",
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60")
        )
        ver_lbl.pack(side="right")

        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="Convert PowerPoint presentations to PDF locally with 100% privacy",
            font=ctk.CTkFont(size=13),
            text_color=("gray45", "gray65")
        )
        subtitle_label.pack(anchor="w", pady=(2, 0))

    def _create_file_list_section(self):
        """Section for selecting and viewing PPTX files."""
        section_frame = ctk.CTkFrame(
            self,
            corner_radius=12,
            border_width=1,
            border_color=("gray85", "gray25")
        )
        section_frame.pack(fill="both", expand=True, padx=24, pady=6)

        # Header bar for file list
        bar_frame = ctk.CTkFrame(section_frame, fg_color="transparent")
        bar_frame.pack(fill="x", padx=14, pady=(12, 6))

        lbl = ctk.CTkLabel(
            bar_frame,
            text="PowerPoint Files (.pptx)",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        lbl.pack(side="left")

        self.file_count_label = ctk.CTkLabel(
            bar_frame,
            text="0 files selected",
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray70")
        )
        self.file_count_label.pack(side="right")

        # Scrollable Frame for Files List
        self.file_scroll_frame = ctk.CTkScrollableFrame(
            section_frame,
            height=190,
            fg_color=("gray95", "gray15"),
            corner_radius=8
        )
        self.file_scroll_frame.pack(fill="both", expand=True, padx=14, pady=4)

        # Placeholder label when no files
        self.empty_label = ctk.CTkLabel(
            self.file_scroll_frame,
            text="📂 No PowerPoint files selected.\n\nClick 'Add Files' below to begin.",
            font=ctk.CTkFont(size=13),
            text_color=("gray50", "gray60")
        )
        self.empty_label.pack(expand=True, pady=45)

        # Button Bar (Add, Remove, Clear)
        btn_frame = ctk.CTkFrame(section_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=14, pady=(8, 12))

        self.btn_add = ctk.CTkButton(
            btn_frame,
            text="➕ Add Files",
            width=130,
            height=34,
            corner_radius=8,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=ACCENT_ORANGE,
            hover_color=ACCENT_HOVER,
            command=self._add_files
        )
        self.btn_add.pack(side="left", padx=(0, 10))

        self.btn_remove = ctk.CTkButton(
            btn_frame,
            text="🗑️ Remove",
            width=105,
            height=34,
            corner_radius=8,
            fg_color="transparent",
            border_width=1,
            border_color=("gray75", "gray35"),
            text_color=("gray10", "gray90"),
            command=self._remove_selected
        )
        self.btn_remove.pack(side="left", padx=(0, 10))

        self.btn_clear = ctk.CTkButton(
            btn_frame,
            text="🧹 Clear All",
            width=105,
            height=34,
            corner_radius=8,
            fg_color="transparent",
            border_width=1,
            border_color=("gray75", "gray35"),
            text_color=("gray10", "gray90"),
            command=self._clear_files
        )
        self.btn_clear.pack(side="left")

    def _create_output_section(self):
        """Section for selecting output folder."""
        section_frame = ctk.CTkFrame(
            self,
            corner_radius=12,
            border_width=1,
            border_color=("gray85", "gray25")
        )
        section_frame.pack(fill="x", padx=24, pady=6)

        lbl = ctk.CTkLabel(
            section_frame,
            text="Output Folder",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        lbl.pack(anchor="w", padx=14, pady=(10, 4))

        folder_frame = ctk.CTkFrame(section_frame, fg_color="transparent")
        folder_frame.pack(fill="x", padx=14, pady=(0, 12))

        self.output_entry = ctk.CTkEntry(
            folder_frame,
            height=36,
            corner_radius=8,
            placeholder_text="Select destination folder for converted PDFs..."
        )
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_browse = ctk.CTkButton(
            folder_frame,
            text="📁 Browse...",
            width=110,
            height=36,
            corner_radius=8,
            fg_color="transparent",
            border_width=1,
            border_color=("gray75", "gray35"),
            text_color=("gray10", "gray90"),
            command=self._browse_output_folder
        )
        self.btn_browse.pack(side="right")

    def _create_settings_section(self):
        """Options & Settings checkboxes."""
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="x", padx=24, pady=4)

        self.var_overwrite = ctk.BooleanVar(value=True)
        self.chk_overwrite = ctk.CTkCheckBox(
            frame,
            text="Overwrite existing PDF files",
            variable=self.var_overwrite,
            hover_color=ACCENT_HOVER,
            fg_color=ACCENT_ORANGE
        )
        self.chk_overwrite.pack(side="left", padx=(0, 24))

        self.var_auto_open = ctk.BooleanVar(value=True)
        self.chk_auto_open = ctk.CTkCheckBox(
            frame,
            text="Open output folder after conversion",
            variable=self.var_auto_open,
            hover_color=ACCENT_HOVER,
            fg_color=ACCENT_ORANGE
        )
        self.chk_auto_open.pack(side="left")

    def _create_action_progress_section(self):
        """Convert button, progress bar, and status messages."""
        section_frame = ctk.CTkFrame(
            self,
            corner_radius=12,
            border_width=1,
            border_color=("gray85", "gray25")
        )
        section_frame.pack(fill="x", padx=24, pady=8)

        # Large Convert Button
        self.btn_convert = ctk.CTkButton(
            section_frame,
            text="🚀 Convert to PDF",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=46,
            corner_radius=10,
            fg_color=ACCENT_ORANGE,
            hover_color=ACCENT_HOVER,
            command=self._start_conversion
        )
        self.btn_convert.pack(fill="x", padx=16, pady=(16, 10))

        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(
            section_frame,
            height=10,
            corner_radius=5,
            progress_color=ACCENT_ORANGE
        )
        self.progress_bar.pack(fill="x", padx=16, pady=(4, 6))
        self.progress_bar.set(0)

        # Status text
        status_frame = ctk.CTkFrame(section_frame, fg_color="transparent")
        status_frame.pack(fill="x", padx=16, pady=(2, 12))

        self.lbl_status = ctk.CTkLabel(
            status_frame,
            text="Status: Ready",
            font=ctk.CTkFont(size=12),
            anchor="w"
        )
        self.lbl_status.pack(side="left", fill="x", expand=True)

        self.lbl_percent = ctk.CTkLabel(
            status_frame,
            text="0%",
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="e"
        )
        self.lbl_percent.pack(side="right")

    def _create_footer(self):
        """Footer with open output folder button and LibreOffice status indicator."""
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.pack(fill="x", padx=24, pady=(0, 14))

        lo_info_frame = ctk.CTkFrame(footer_frame, fg_color="transparent")
        lo_info_frame.pack(side="left")

        self.lbl_lo_status = ctk.CTkLabel(
            lo_info_frame,
            text="Checking LibreOffice...",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        self.lbl_lo_status.pack(side="left", padx=(0, 10))

        self.btn_download_lo = ctk.CTkButton(
            lo_info_frame,
            text="🌐 Download LibreOffice",
            width=150,
            height=28,
            corner_radius=6,
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=ACCENT_BLUE,
            hover_color="#01579B",
            command=self._open_libreoffice_download
        )
        self.btn_download_lo.pack(side="left", padx=(0, 6))

        self.btn_locate_lo = ctk.CTkButton(
            lo_info_frame,
            text="🔍 Locate soffice.exe",
            width=135,
            height=28,
            corner_radius=6,
            font=ctk.CTkFont(size=11),
            fg_color="transparent",
            border_width=1,
            border_color=("gray75", "gray35"),
            text_color=("gray10", "gray90"),
            command=self._locate_libreoffice_binary
        )
        self.btn_locate_lo.pack(side="left")

        self.btn_open_folder = ctk.CTkButton(
            footer_frame,
            text="📂 Open Output Folder",
            width=160,
            height=32,
            corner_radius=8,
            fg_color="transparent",
            border_width=1,
            border_color=("gray75", "gray35"),
            text_color=("gray10", "gray90"),
            command=self._on_open_output_folder
        )
        self.btn_open_folder.pack(side="right")

    # --- UI Logic & Handlers ---

    def _open_libreoffice_download(self):
        """Open LibreOffice official download page in browser."""
        import webbrowser
        webbrowser.open("https://www.libreoffice.org/download/download/")

    def _locate_libreoffice_binary(self):
        """Browse to custom soffice.exe location."""
        filename = filedialog.askopenfilename(
            title="Select LibreOffice soffice.exe",
            filetypes=[("LibreOffice Executable", "soffice.exe;soffice"), ("All Files", "*.*")]
        )
        if filename:
            if self.converter.set_libreoffice_path(filename):
                self._check_libreoffice_on_startup()
                messagebox.showinfo("LibreOffice Located", f"Successfully set LibreOffice path:\n{filename}")
            else:
                messagebox.showerror("Invalid Executable", "The selected file is not a valid soffice.exe executable.")

    def _check_libreoffice_on_startup(self):
        """Check for LibreOffice on app startup or re-detect."""
        if not self.converter.is_ready():
            self.converter.find_and_set_libreoffice() if hasattr(self.converter, 'find_and_set_libreoffice') else None

        lo_path = self.converter.libreoffice_path
        if lo_path:
            self.lbl_lo_status.configure(
                text=f"✓ LibreOffice: {lo_path.name}",
                text_color=("green", "#4CAF50")
            )
            self.btn_download_lo.pack_forget()
            self.btn_locate_lo.pack_forget()
        else:
            self.lbl_lo_status.configure(
                text="⚠️ LibreOffice not found",
                text_color=("red", "#F44336")
            )
            self.btn_download_lo.pack(side="left", padx=(0, 6))
            self.btn_locate_lo.pack(side="left")

    def _add_files(self):
        """Open file dialog to add PPTX files."""
        filenames = filedialog.askopenfilenames(
            title="Select PowerPoint Presentations",
            filetypes=[("PowerPoint Presentations", "*.pptx"), ("All Files", "*.*")]
        )
        if not filenames:
            return

        added_any = False
        for fname in filenames:
            path = Path(fname)
            if is_valid_pptx(path) and path not in self.selected_files:
                self.selected_files.append(path)
                added_any = True

        if added_any:
            current_output = self.output_entry.get().strip()
            if not current_output and self.selected_files:
                first_dir = self.selected_files[0].parent
                self.output_entry.delete(0, "end")
                self.output_entry.insert(0, str(first_dir))

            self._update_file_list_display()

    def _remove_selected(self):
        """Remove highlighted/selected items from list."""
        if self.selected_files:
            self.selected_files.pop()
            self._update_file_list_display()

    def _clear_files(self):
        """Clear all selected files."""
        self.selected_files.clear()
        self._update_file_list_display()

    def _update_file_list_display(self):
        """Re-render the scrollable list of selected files as styled cards."""
        for child in self.file_scroll_frame.winfo_children():
            child.destroy()

        count = len(self.selected_files)
        self.file_count_label.configure(text=f"{count} file{'s' if count != 1 else ''} selected")

        if not self.selected_files:
            self.empty_label = ctk.CTkLabel(
                self.file_scroll_frame,
                text="📂 No PowerPoint files selected.\n\nClick 'Add Files' below to begin.",
                font=ctk.CTkFont(size=13),
                text_color=("gray50", "gray60")
            )
            self.empty_label.pack(expand=True, pady=45)
            return

        for idx, file_path in enumerate(self.selected_files, start=1):
            item_frame = ctk.CTkFrame(
                self.file_scroll_frame,
                fg_color=("gray90", "gray22"),
                corner_radius=8,
                height=42
            )
            item_frame.pack(fill="x", pady=3, padx=2)

            num_badge = ctk.CTkFrame(item_frame, fg_color=ACCENT_ORANGE, corner_radius=6, width=28, height=24)
            num_badge.pack(side="left", padx=(8, 4))
            num_badge.pack_propagate(False)

            num_lbl = ctk.CTkLabel(
                num_badge,
                text=str(idx),
                font=ctk.CTkFont(size=11, weight="bold"),
                text_color="white"
            )
            num_lbl.pack(expand=True)

            # PPTX icon tag
            tag_lbl = ctk.CTkLabel(
                item_frame,
                text="📄",
                font=ctk.CTkFont(size=14)
            )
            tag_lbl.pack(side="left", padx=(4, 6))

            name_lbl = ctk.CTkLabel(
                item_frame,
                text=file_path.name,
                font=ctk.CTkFont(size=12, weight="bold"),
                anchor="w"
            )
            name_lbl.pack(side="left", padx=2)

            path_lbl = ctk.CTkLabel(
                item_frame,
                text=f"({file_path.parent})",
                font=ctk.CTkFont(size=11),
                text_color=("gray50", "gray60"),
                anchor="w"
            )
            path_lbl.pack(side="left", padx=8, fill="x", expand=True)

            del_btn = ctk.CTkButton(
                item_frame,
                text="✕",
                width=28,
                height=28,
                corner_radius=14,
                fg_color="transparent",
                hover_color=("gray80", "gray35"),
                text_color=("gray30", "gray70"),
                command=lambda p=file_path: self._remove_specific_file(p)
            )
            del_btn.pack(side="right", padx=6)

    def _remove_specific_file(self, file_path: Path):
        """Remove a specific file from selected list."""
        if file_path in self.selected_files:
            self.selected_files.remove(file_path)
            self._update_file_list_display()

    def _browse_output_folder(self):
        """Open directory dialog for output folder."""
        initial_dir = self.output_entry.get().strip() or os.path.expanduser("~")
        folder = filedialog.askdirectory(title="Select Output Folder", initialdir=initial_dir)
        if folder:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, folder)

    def _on_open_output_folder(self):
        """Open currently configured output directory."""
        folder = self.output_entry.get().strip()
        if not folder:
            messagebox.showwarning("No Output Folder", "Please specify an output folder first.")
            return
        success, err = open_output_folder(folder)
        if not success:
            messagebox.showerror("Error Opening Folder", err)

    def _set_ui_converting_state(self, converting: bool):
        """Enable or disable interactive widgets during conversion."""
        self.is_converting = converting
        state = "disabled" if converting else "normal"

        self.btn_add.configure(state=state)
        self.btn_remove.configure(state=state)
        self.btn_clear.configure(state=state)
        self.btn_browse.configure(state=state)
        self.btn_convert.configure(state=state)
        self.chk_overwrite.configure(state=state)
        self.chk_auto_open.configure(state=state)
        self.output_entry.configure(state=state)

        if converting:
            self.btn_convert.configure(text="⌛ Converting...", fg_color="gray")
        else:
            self.btn_convert.configure(text="🚀 Convert to PDF", fg_color=ACCENT_ORANGE)

    # --- Asynchronous Conversion Worker ---

    def _start_conversion(self):
        """Validate inputs and launch conversion thread."""
        if self.is_converting:
            return

        if not self.selected_files:
            messagebox.showwarning(
                "No Files Selected",
                "Please add at least one PowerPoint (.pptx) file to convert."
            )
            return

        output_dir = self.output_entry.get().strip()
        if not output_dir:
            messagebox.showwarning(
                "No Output Folder",
                "Please select an output folder before converting."
            )
            return

        if not self.converter.is_ready():
            messagebox.showerror(
                "LibreOffice Not Found",
                "LibreOffice was not found on your computer.\n\n"
                "LibreOffice is required to render and convert PowerPoint presentations locally.\n"
                "Please install LibreOffice or click 'Download LibreOffice' below."
            )
            return

        self._set_ui_converting_state(True)
        self.progress_bar.set(0)
        self.lbl_percent.configure(text="0%")
        self.lbl_status.configure(text="Status: Preparing conversion...")

        thread = threading.Thread(
            target=self._run_conversion_worker,
            args=(list(self.selected_files), output_dir, self.var_overwrite.get()),
            daemon=True
        )
        thread.start()

    def _run_conversion_worker(self, files: list[Path], output_dir: str, overwrite: bool):
        """Background thread target performing batch conversion."""
        def progress_cb(current_idx: int, total_count: int, filename: str, status_msg: str):
            percent = (current_idx - 1) / total_count
            self.after(0, self._update_progress_ui, percent, f"Status: {status_msg}", f"{int(percent * 100)}%")

        result: BatchConversionResult = self.converter.convert_batch(
            input_files=files,
            output_folder=output_dir,
            overwrite=overwrite,
            progress_callback=progress_cb
        )

        self.after(0, self._on_conversion_complete, result, output_dir)

    def _update_progress_ui(self, progress_val: float, status_text: str, percent_text: str):
        """Update progress bar and status labels on main thread."""
        self.progress_bar.set(progress_val)
        self.lbl_status.configure(text=status_text)
        self.lbl_percent.configure(text=percent_text)

    def _on_conversion_complete(self, result: BatchConversionResult, output_dir: str):
        """Completion callback executed on the main GUI thread."""
        self.progress_bar.set(1.0)
        self.lbl_percent.configure(text="100%")
        self._set_ui_converting_state(False)

        if result.failed_count == 0:
            self.lbl_status.configure(text=f"Status: Complete! Converted {result.successful_count} file(s) in {result.total_duration}s.")
            message = (
                f"Successfully converted {result.successful_count} presentation(s) to PDF!\n\n"
                f"Time taken: {result.total_duration} seconds\n"
                f"Output folder: {output_dir}"
            )
            messagebox.showinfo("Conversion Complete", message)
        else:
            self.lbl_status.configure(
                text=f"Status: Finished with errors ({result.successful_count} succeeded, {result.failed_count} failed)."
            )
            failed_details = "\n".join([f"• {r.input_path.name}: {r.error_message}" for r in result.results if not r.success])
            message = (
                f"Conversion completed with errors.\n\n"
                f"Successfully converted: {result.successful_count}\n"
                f"Failed: {result.failed_count}\n\n"
                f"Failed Files:\n{failed_details}\n\n"
                f"See logs for full technical details."
            )
            messagebox.showwarning("Conversion Finished with Errors", message)

        if self.var_auto_open.get() and result.successful_count > 0:
            open_output_folder(output_dir)
