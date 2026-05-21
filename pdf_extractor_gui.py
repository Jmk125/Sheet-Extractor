import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, Toplevel, Scrollbar, IntVar, Frame, Canvas, Label, Button, Entry, Checkbutton
from PIL import Image, ImageTk
import os
import re


class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Sheet Extractor")
        self.root.geometry("1200x780")
        self.root.configure(bg="#f3f6fb")

        self.pdf_path = ""
        self.page_number = 1
        self.rect_coords_number = None
        self.rect_coords_title = None
        self.pdf_document = None
        self.scale_factor = 1.0

        self.draw_phase = None
        self.pending_number_box = None
        self.retry_indices = []
        self.active_retry_indices = []
        self.scan_start_page = 1

        self.h_line = None
        self.v_line = None

        self.status_text = tk.StringVar(value="Upload a PDF to begin.")
        self.extraction_mode = "drawings"

        self.button_frame = Frame(root, bg="#ffffff", padx=10, pady=10)
        self.button_frame.pack(fill=tk.X, padx=10, pady=(10, 4))

        self.mode_button = Button(
            self.button_frame,
            text="Mode: Drawings",
            command=self.toggle_mode,
            bg="#0ea5e9",
            fg="white",
            activebackground="#0284c7",
            activeforeground="white",
            relief=tk.FLAT,
            padx=14,
            pady=8,
        )
        self.mode_button.pack(side=tk.LEFT, padx=(0, 8))

        self.upload_button = Button(
            self.button_frame,
            text="Upload PDF",
            command=self.upload_pdf,
            bg="#2563eb",
            fg="white",
            activebackground="#1d4ed8",
            activeforeground="white",
            relief=tk.FLAT,
            padx=14,
            pady=8,
        )
        self.upload_button.pack(side=tk.LEFT, padx=(0, 8))

        self.prev_button = Button(
            self.button_frame,
            text="Previous Drawing",
            command=self.previous_page,
            bg="#e5e7eb",
            fg="#111827",
            relief=tk.FLAT,
            padx=14,
            pady=8,
        )
        self.prev_button.pack(side=tk.LEFT)

        self.next_button = Button(
            self.button_frame,
            text="Next Drawing",
            command=self.next_drawing,
            bg="#e5e7eb",
            fg="#111827",
            relief=tk.FLAT,
            padx=14,
            pady=8,
        )
        self.next_button.pack(side=tk.LEFT, padx=(8, 0))

        self.status_label = Label(
            root,
            textvariable=self.status_text,
            anchor="w",
            bg="#eff6ff",
            fg="#1e3a8a",
            font=("Segoe UI", 11, "bold"),
            padx=12,
            pady=8,
        )
        self.status_label.pack(fill=tk.X, padx=10, pady=(0, 8))

        self.canvas_frame = Frame(root, bg="#dbeafe", bd=1, relief=tk.SOLID)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.canvas = Canvas(self.canvas_frame, bg="#ffffff", highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar_y = Scrollbar(self.canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.scrollbar_x = Scrollbar(root, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X, padx=10)

        self.canvas.config(yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set)

        self.canvas.bind("<Button-1>", self.start_draw)
        self.canvas.bind("<B1-Motion>", self.draw_rect)
        self.canvas.bind("<ButtonRelease-1>", self.end_draw)
        self.canvas.bind("<Button-3>", self.start_pan)
        self.canvas.bind("<B3-Motion>", self.pan_canvas)
        self.canvas.bind("<Motion>", self.update_crosshair)
        self.canvas.bind("<Enter>", self.show_crosshair)
        self.canvas.bind("<Leave>", self.hide_crosshair)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)

    def set_status(self, text):
        self.status_text.set(text)


    def _create_styled_popup(self, title, geometry="440x220"):
        popup = Toplevel(self.root)
        popup.title(title)
        popup.geometry(geometry)
        popup.configure(bg="#f8fafc")
        popup.transient(self.root)
        popup.grab_set()
        return popup

    def _popup_content_frame(self, popup):
        frame = Frame(popup, bg="#f8fafc", padx=16, pady=14)
        frame.pack(fill=tk.BOTH, expand=True)
        return frame

    def _styled_button(self, parent, text, command, primary=False):
        if primary:
            return Button(parent, text=text, command=command, bg="#2563eb", fg="white", activebackground="#1d4ed8", activeforeground="white", relief=tk.FLAT, padx=12, pady=6)
        return Button(parent, text=text, command=command, bg="#e5e7eb", fg="#111827", activebackground="#d1d5db", relief=tk.FLAT, padx=12, pady=6)

    def _styled_panel_button(self, parent, text, command, primary=False, width=None):
        if primary:
            return Button(parent, text=text, command=command, bg="#2563eb", fg="white", activebackground="#1d4ed8", activeforeground="white", relief=tk.FLAT, padx=10, pady=6, width=width)
        return Button(parent, text=text, command=command, bg="#e2e8f0", fg="#0f172a", activebackground="#cbd5e1", relief=tk.FLAT, padx=10, pady=6, width=width)

    def show_crosshair(self, event):
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        self.create_crosshair(x, y)

    def hide_crosshair(self, event):
        if self.h_line:
            self.canvas.delete(self.h_line)
            self.h_line = None
        if self.v_line:
            self.canvas.delete(self.v_line)
            self.v_line = None

    def create_crosshair(self, x, y):
        bbox = self.canvas.bbox(tk.ALL)
        if not bbox:
            width = self.canvas.winfo_width()
            height = self.canvas.winfo_height()
            left, top = 0, 0
            right, bottom = width, height
        else:
            left, top, right, bottom = bbox

        if self.h_line:
            self.canvas.delete(self.h_line)
        if self.v_line:
            self.canvas.delete(self.v_line)

        self.h_line = self.canvas.create_line(left, y, right, y, fill="#2563eb", dash=(4, 4))
        self.v_line = self.canvas.create_line(x, top, x, bottom, fill="#2563eb", dash=(4, 4))
        self.canvas.tag_raise(self.h_line)
        self.canvas.tag_raise(self.v_line)
        if hasattr(self, "rect") and self.rect:
            self.canvas.tag_raise(self.rect)

    def update_crosshair(self, event):
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        self.create_crosshair(x, y)

    def start_pan(self, event):
        self.canvas.scan_mark(event.x, event.y)
        self.set_status("Panning drawing view...")

    def pan_canvas(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def toggle_mode(self):
        if self.extraction_mode == "drawings":
            self.extraction_mode = "specs"
            self.mode_button.config(text="Mode: Specs")
            self.prev_button.config(text="Previous Page")
            self.next_button.config(text="Next Page")
            self.set_status("Specs mode enabled. Upload a large specs PDF to begin.")
        else:
            self.extraction_mode = "drawings"
            self.mode_button.config(text="Mode: Drawings")
            self.prev_button.config(text="Previous Drawing")
            self.next_button.config(text="Next Drawing")
            self.set_status("Drawings mode enabled. Upload a PDF to begin.")

    def upload_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.pdf_path:
            self.page_number = 1
            self.rect_coords_number = None
            self.rect_coords_title = None
            self.draw_phase = "number"
            self.scan_start_page = 1
            self.load_first_page()
            if self.extraction_mode == "specs":
                self.set_status("Go to the first technical spec sheet, then draw a box around the spec number")
            else:
                self.set_status("Draw a box around the sheet number")

    def next_drawing(self):
        if self.pdf_document and self.page_number < len(self.pdf_document):
            self.page_number += 1
            self.display_page(self.page_number - 1)

    def previous_page(self):
        if self.pdf_document and self.page_number > 1:
            self.page_number -= 1
            self.display_page(self.page_number - 1)

    def load_first_page(self):
        self.pdf_document = fitz.open(self.pdf_path)
        self.display_page(self.page_number - 1)

    def display_page(self, page_index):
        page = self.pdf_document.load_page(page_index)
        pix = page.get_pixmap(matrix=fitz.Matrix(self.scale_factor, self.scale_factor))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        self.img_tk = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.img_tk)
        self.canvas.config(scrollregion=self.canvas.bbox(tk.ALL))

    def start_draw(self, event):
        if not self.pdf_document or not self.draw_phase:
            return
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="#dc2626", width=2)
        self.update_crosshair(event)

    def draw_rect(self, event):
        if not self.pdf_document or not hasattr(self, "rect"):
            return
        current_x = self.canvas.canvasx(event.x)
        current_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, current_x, current_y)
        self.update_crosshair(event)
        self.canvas.tag_raise(self.rect)

    def end_draw(self, event):
        if not self.pdf_document or not self.draw_phase:
            return
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        new_box = (self.start_x, self.start_y, end_x, end_y)

        if self.draw_phase == "number":
            if not self.active_retry_indices:
                self.scan_start_page = self.page_number
            self.pending_number_box = new_box
            scaled_rect_coords = [coord / self.scale_factor for coord in new_box]
            extracted_text_number = self.extract_text_from_box(self.page_number, scaled_rect_coords)
            if extracted_text_number:
                self.show_confirmation_popup_number(extracted_text_number)
            else:
                self.show_notice_popup("No Text", "No text found in the selected area.")
        elif self.draw_phase == "title":
            self.rect_coords_title = new_box
            scaled_rect_coords = [coord / self.scale_factor for coord in new_box]
            extracted_text_title = self.extract_text_from_box(self.page_number, scaled_rect_coords)
            if extracted_text_title:
                extracted_text_title = " ".join(extracted_text_title.split())
                self.show_confirmation_popup_title(extracted_text_title)
            else:
                self.show_notice_popup("No Text", "No text found in the selected area.")

    def extract_text_from_box(self, page_number, box):
        try:
            page_count = self.pdf_document.page_count
            if page_number < 1 or page_number > page_count:
                return ""
            page = self.pdf_document.load_page(page_number - 1)
            rect = fitz.Rect(box)
            return page.get_text("text", clip=rect).strip()
        except Exception:
            return ""

    def show_confirmation_popup_number(self, extracted_text_number):
        confirmation_popup = self._create_styled_popup("Confirm Sheet Number")
        content_frame = self._popup_content_frame(confirmation_popup)

        number_label = "Confirm Spec Number" if self.extraction_mode == "specs" else "Confirm Sheet Number"
        title_prompt = "Confirm & Draw Spec Name Box" if self.extraction_mode == "specs" else "Confirm & Draw Title Box"
        display_number = self.normalize_spec_number(extracted_text_number) if self.extraction_mode == "specs" else extracted_text_number

        Label(
            content_frame,
            text=number_label,
            font=("Segoe UI", 12, "bold"),
            bg="#f8fafc",
            fg="#0f172a",
        ).pack(anchor="w", pady=(0, 8))
        Label(content_frame, text=f"Extracted text:\n{display_number}", bg="#f8fafc", fg="#334155", justify=tk.LEFT, wraplength=390).pack(anchor="w", pady=(0, 14))

        button_frame = Frame(content_frame, bg="#f8fafc")
        button_frame.pack(anchor="e")

        self._styled_button(button_frame, "Cancel", confirmation_popup.destroy).pack(side=tk.LEFT, padx=6)
        self._styled_button(button_frame, "Skip Title", lambda: [self.skip_title_drawing(), confirmation_popup.destroy()]).pack(side=tk.LEFT, padx=6)
        self._styled_button(
            button_frame,
            title_prompt,
            lambda: [self.activate_title_box_drawing(), confirmation_popup.destroy()],
            primary=True,
        ).pack(side=tk.LEFT, padx=6)

    def skip_title_drawing(self):
        self.rect_coords_number = self.pending_number_box
        self.rect_coords_title = None
        self.draw_phase = None
        self.set_status("Extracting sheets...")
        self.process_all_pages()

    def activate_title_box_drawing(self):
        self.rect_coords_number = self.pending_number_box
        self.draw_phase = "title"
        self.set_status("Draw a box around the spec name" if self.extraction_mode == "specs" else "Draw a box around the sheet title")

    def show_confirmation_popup_title(self, extracted_text_title):
        confirmation_popup = self._create_styled_popup("Confirm Name", geometry="440x210")
        content_frame = self._popup_content_frame(confirmation_popup)

        label_text = "Confirm Spec Name" if self.extraction_mode == "specs" else "Confirm Sheet Title"
        Label(content_frame, text=label_text, font=("Segoe UI", 12, "bold"), bg="#f8fafc", fg="#0f172a").pack(anchor="w", pady=(0, 8))
        Label(content_frame, text=f"Extracted text:\n{extracted_text_title}", bg="#f8fafc", fg="#334155", justify=tk.LEFT, wraplength=390).pack(anchor="w", pady=(0, 14))

        button_frame = Frame(content_frame, bg="#f8fafc")
        button_frame.pack(anchor="e")

        self._styled_button(button_frame, "Cancel", confirmation_popup.destroy).pack(side=tk.LEFT, padx=6)
        self._styled_button(
            button_frame,
            "Confirm",
            lambda: [self.finish_template_and_process(), confirmation_popup.destroy()],
            primary=True,
        ).pack(side=tk.LEFT, padx=6)

    def finish_template_and_process(self):
        if self.active_retry_indices:
            self.draw_phase = None
            self.apply_retry_to_checked_rows()
            return
        self.draw_phase = None
        self.set_status("Extracting sheets...")
        self.process_all_pages()

    def process_all_pages(self):
        self.sheet_numbers_titles = []
        page_count = self.pdf_document.page_count
        start_page = self.scan_start_page if self.extraction_mode == "specs" else 1
        start_page = max(1, min(start_page, page_count))
        for page_num in range(start_page, page_count + 1):
            scaled_rect_coords_number = [coord / self.scale_factor for coord in self.rect_coords_number]
            text_number = self.extract_text_from_box(page_num, scaled_rect_coords_number)
            if self.rect_coords_title:
                scaled_rect_coords_title = [coord / self.scale_factor for coord in self.rect_coords_title]
                text_title = self.extract_text_from_box(page_num, scaled_rect_coords_title)
            else:
                text_title = ""
            self.sheet_numbers_titles.append((page_num, text_number, text_title))

        if self.extraction_mode == "specs":
            self.build_specs_summary()
            self.show_specs_selection()
        else:
            self.show_sheet_selection()
        self.set_status("Review results. You can retry checked rows with new boxes.")

    def build_specs_summary(self):
        grouped_specs = {}
        missing_pages = []
        for page_num, text_number, text_title in self.sheet_numbers_titles:
            spec_number = self.normalize_spec_number(text_number)
            spec_name = text_title.strip()
            if not spec_number and not spec_name:
                missing_pages.append(page_num)
                continue
            if spec_number not in grouped_specs:
                grouped_specs[spec_number] = {"name": spec_name, "pages": []}
            elif not grouped_specs[spec_number]["name"] and spec_name:
                grouped_specs[spec_number]["name"] = spec_name
            grouped_specs[spec_number]["pages"].append(page_num)
        self.spec_groups = grouped_specs
        self.spec_missing_pages = missing_pages

    def normalize_spec_number(self, raw_text):
        text = (raw_text or "").upper()
        text = re.sub(r"[–—−]", "-", text)

        # Keep extraction conservative: find the first true 6-digit section number,
        # then ignore anything after it (e.g., "-1", "- 2", suffix text).
        spaced_match = re.search(r"(?<!\d)(\d(?:\s*\d){5})(?!\d)", text)
        if spaced_match:
            return re.sub(r"\s+", "", spaced_match.group(1))

        compact = re.sub(r"\s+", "", text)
        compact = re.sub(r"(?<=\d)[O](?=\d)", "0", compact)
        compact = re.sub(r"(?<!\d)[O](?=\d{5}(?!\d))", "0", compact)
        compact = re.sub(r"(?<=\d)[IL](?=\d)", "1", compact)

        match = re.search(r"(?<!\d)(\d{6})(?!\d)", compact)
        if match:
            return match.group(1)

        return ""

    def show_sheet_selection(self):
        if hasattr(self, "selection_window") and self.selection_window.winfo_exists():
            self.selection_window.destroy()

        selection_window = Toplevel(self.root)
        selection_window.title("Select Sheets to Extract")
        selection_window.geometry("900x620")
        selection_window.configure(bg="#f3f6fb")
        self.selection_window = selection_window

        toolbar_frame = Frame(selection_window, bg="#ffffff", padx=12, pady=10)
        toolbar_frame.pack(side=tk.TOP, pady=(10, 6), padx=10, fill=tk.X)

        header_text = "Review extracted specs" if self.extraction_mode == "specs" else "Review extracted sheets"
        Label(toolbar_frame, text=header_text, bg="#ffffff", fg="#0f172a", font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT, padx=(0, 14))

        button_frame_top = Frame(toolbar_frame, bg="#ffffff")
        button_frame_top.pack(side=tk.LEFT)

        self._styled_panel_button(button_frame_top, "Check All", self.check_all).pack(side=tk.LEFT, padx=4)
        self._styled_panel_button(button_frame_top, "Uncheck All", self.uncheck_all).pack(side=tk.LEFT, padx=4)
        if self.extraction_mode == "drawings":
            self._styled_panel_button(button_frame_top, "A Drawings", lambda: self.check_drawings_by_letter("A")).pack(side=tk.LEFT, padx=4)
            self._styled_panel_button(button_frame_top, "C Drawings", lambda: self.check_drawings_by_letter("C")).pack(side=tk.LEFT, padx=4)
            self._styled_panel_button(button_frame_top, "S Drawings", lambda: self.check_drawings_by_letter("S")).pack(side=tk.LEFT, padx=4)
        self._styled_panel_button(toolbar_frame, "Retry Checked Rows", self.start_retry_for_checked, primary=True).pack(side=tk.RIGHT, padx=4)

        header_frame = Frame(selection_window, bg="#eff6ff", padx=12, pady=8)
        header_frame.pack(fill=tk.X, padx=10, pady=(0, 6))
        Label(header_frame, text="Select rows to export or retry. You can edit number/title text directly.", bg="#eff6ff", fg="#1e3a8a", font=("Segoe UI", 10, "bold")).pack(anchor="w")

        checklist_frame = Frame(selection_window, bg="#ffffff", bd=1, relief=tk.SOLID)
        checklist_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 8))

        self.checklist_canvas = Canvas(checklist_frame, bg="#ffffff", highlightthickness=0)
        self.checklist_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = Scrollbar(checklist_frame, orient=tk.VERTICAL, command=self.checklist_canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.checklist_canvas.config(yscrollcommand=scrollbar.set)

        checklist_inner_frame = Frame(self.checklist_canvas, bg="#ffffff")
        self.checklist_canvas.create_window((0, 0), window=checklist_inner_frame, anchor=tk.NW)
        checklist_inner_frame.bind("<Configure>", lambda e: self.checklist_canvas.config(scrollregion=self.checklist_canvas.bbox(tk.ALL)))
        self.checklist_canvas.bind_all("<MouseWheel>", self._on_checklist_mousewheel)
        self.checklist_canvas.bind_all("<MouseWheel>", self._on_checklist_mousewheel)

        self.check_vars = []
        self.entries_number = []
        self.entries_title = []

        for page_num, text_number, text_title in self.sheet_numbers_titles:
            var = IntVar()
            frame = Frame(checklist_inner_frame, bg="#ffffff")
            frame.pack(anchor=tk.W, fill=tk.X)

            Checkbutton(frame, variable=var, bg="#ffffff", activebackground="#ffffff").pack(side=tk.LEFT, padx=5)
            self.check_vars.append(var)

            Label(frame, text=f"Pg {page_num}", width=8, anchor="w", bg="#ffffff", fg="#334155").pack(side=tk.LEFT)

            entry_number = Entry(frame)
            entry_number.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            entry_number.insert(0, text_number)
            self.entries_number.append(entry_number)

            entry_title = Entry(frame, width=44)
            entry_title.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            entry_title.insert(0, text_title)
            self.entries_title.append(entry_title)

        button_frame_bottom = Frame(selection_window, bg="#ffffff", padx=12, pady=10)
        button_frame_bottom.pack(side=tk.BOTTOM, pady=(0, 10), padx=10, fill=tk.X)

        if self.extraction_mode == "specs":
            self._styled_panel_button(button_frame_bottom, "Preview Checked", self.preview_checked_pages, width=20).pack(side=tk.LEFT, padx=5)
            self._styled_panel_button(button_frame_bottom, "Save Specs by Number + Name", lambda: self.save_specs_grouped(), width=32, primary=True).pack(side=tk.LEFT, padx=5)
        else:
            self._styled_panel_button(button_frame_bottom, "Save Sheets with Number", lambda: self.save_sheets(False), width=25).pack(side=tk.LEFT, padx=5)
            self._styled_panel_button(
                button_frame_bottom,
                text="Save Sheets with Number and Title",
                width=30,
                command=lambda: self.save_sheets(True),
                primary=True,
            ).pack(side=tk.LEFT, padx=5)

    def show_specs_selection(self):
        if hasattr(self, "selection_window") and self.selection_window.winfo_exists():
            self.selection_window.destroy()

        selection_window = Toplevel(self.root)
        selection_window.title("Select Specs to Extract")
        selection_window.geometry("900x620")
        selection_window.configure(bg="#f3f6fb")
        self.selection_window = selection_window

        toolbar_frame = Frame(selection_window, bg="#ffffff", padx=12, pady=10)
        toolbar_frame.pack(side=tk.TOP, pady=(10, 6), padx=10, fill=tk.X)
        Label(toolbar_frame, text="Review extracted specs", bg="#ffffff", fg="#0f172a", font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT, padx=(0, 14))

        button_frame_top = Frame(toolbar_frame, bg="#ffffff")
        button_frame_top.pack(side=tk.LEFT)
        self._styled_panel_button(button_frame_top, "Check All", self.check_all).pack(side=tk.LEFT, padx=4)
        self._styled_panel_button(button_frame_top, "Uncheck All", self.uncheck_all).pack(side=tk.LEFT, padx=4)
        self._styled_panel_button(button_frame_top, "Preview Checked", self.preview_checked_pages).pack(side=tk.LEFT, padx=4)

        checklist_frame = Frame(selection_window, bg="#ffffff", bd=1, relief=tk.SOLID)
        checklist_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 8))
        self.checklist_canvas = Canvas(checklist_frame, bg="#ffffff", highlightthickness=0)
        self.checklist_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = Scrollbar(checklist_frame, orient=tk.VERTICAL, command=self.checklist_canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.checklist_canvas.config(yscrollcommand=scrollbar.set)
        checklist_inner_frame = Frame(self.checklist_canvas, bg="#ffffff")
        self.checklist_canvas.create_window((0, 0), window=checklist_inner_frame, anchor=tk.NW)
        checklist_inner_frame.bind("<Configure>", lambda e: self.checklist_canvas.config(scrollregion=self.checklist_canvas.bbox(tk.ALL)))

        self.check_vars = []
        self.spec_entries_number = []
        self.spec_entries_title = []
        self.spec_group_items = list(self.spec_groups.items())

        for spec_number, group_data in self.spec_group_items:
            spec_name = group_data["name"]
            pages = group_data["pages"]
            var = IntVar(value=1)
            frame = Frame(checklist_inner_frame, bg="#ffffff")
            frame.pack(anchor=tk.W, fill=tk.X, pady=1)
            Checkbutton(frame, variable=var, bg="#ffffff", activebackground="#ffffff").pack(side=tk.LEFT, padx=5)
            self.check_vars.append(var)
            Label(frame, text=f"{len(pages)} page(s)", width=12, anchor="w", bg="#ffffff", fg="#334155").pack(side=tk.LEFT, padx=(0, 4))
            entry_number = Entry(frame, width=20)
            entry_number.pack(side=tk.LEFT, padx=5)
            entry_number.insert(0, spec_number)
            self.spec_entries_number.append(entry_number)
            entry_title = Entry(frame, width=72)
            entry_title.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            entry_title.insert(0, spec_name)
            self.spec_entries_title.append(entry_title)

        self.missing_retry_vars = []
        self.missing_retry_pages = self.spec_missing_pages[:]
        for page_num in self.missing_retry_pages:
            var = IntVar()
            frame = Frame(checklist_inner_frame, bg="#fff7ed")
            frame.pack(anchor=tk.W, fill=tk.X, pady=1)
            Checkbutton(frame, variable=var, bg="#fff7ed", activebackground="#fff7ed").pack(side=tk.LEFT, padx=5)
            self.missing_retry_vars.append(var)
            Label(frame, text="Missing", width=12, anchor="w", bg="#fff7ed", fg="#9a3412").pack(side=tk.LEFT, padx=(0, 4))
            Label(frame, text="NO_SPEC_FOUND", width=20, anchor="w", bg="#fff7ed", fg="#9a3412").pack(side=tk.LEFT, padx=5)
            Label(frame, text=f"Pg {page_num}", anchor="w", bg="#fff7ed", fg="#9a3412").pack(side=tk.LEFT, padx=5)

        button_frame_bottom = Frame(selection_window, bg="#ffffff", padx=12, pady=10)
        button_frame_bottom.pack(side=tk.BOTTOM, pady=(0, 10), padx=10, fill=tk.X)
        self._styled_panel_button(button_frame_bottom, "Retry Checked Missing Pages", self.start_retry_missing_specs, width=28).pack(side=tk.LEFT, padx=5)
        self._styled_panel_button(button_frame_bottom, "Save Specs by Number + Name", self.save_specs_grouped, width=32, primary=True).pack(side=tk.LEFT, padx=5)

    def start_retry_for_checked(self):
        self.retry_indices = [i for i, var in enumerate(self.check_vars) if var.get() == 1]
        if not self.retry_indices:
            self.show_notice_popup("No Rows Selected", "Check one or more rows to retry.")
            return
        self.active_retry_indices = self.retry_indices[:]
        first_page = self.sheet_numbers_titles[self.active_retry_indices[0]][0]
        self.page_number = first_page
        self.display_page(first_page - 1)
        self.pending_number_box = None
        self.rect_coords_title = None
        self.draw_phase = "number"
        self.set_status(f"Retry mode: jumped to page {first_page}. Draw a box around the sheet number")
        self.show_notice_popup("Retry Checked Rows", "Draw number box, confirm, then draw title box. New boxes will be applied to all checked rows.")

    def start_retry_missing_specs(self):
        if not hasattr(self, "missing_retry_vars"):
            return
        selected_pages = [self.missing_retry_pages[i] for i, var in enumerate(self.missing_retry_vars) if var.get() == 1]
        if not selected_pages:
            self.show_notice_popup("No Pages Selected", "Check one or more missing pages to retry.")
            return
        page_to_index = {page_num: idx for idx, (page_num, _, _) in enumerate(self.sheet_numbers_titles)}
        self.active_retry_indices = [page_to_index[p] for p in selected_pages if p in page_to_index]
        if not self.active_retry_indices:
            self.show_notice_popup("Retry Error", "Could not map missing pages for retry.")
            return
        first_page = min(selected_pages)
        self.page_number = first_page
        self.display_page(first_page - 1)
        self.pending_number_box = None
        self.rect_coords_title = None
        self.draw_phase = "number"
        self.set_status(f"Retry mode: jumped to page {first_page}. Draw a box around the spec number")
        self.show_notice_popup("Retry Missing Pages", "Draw number box, confirm, then draw name box. New boxes will be applied to selected missing pages.")

    def apply_retry_to_checked_rows(self):
        if not self.active_retry_indices:
            return
        for index in self.active_retry_indices:
            page_num, _, _ = self.sheet_numbers_titles[index]
            scaled_rect_coords_number = [coord / self.scale_factor for coord in self.rect_coords_number]
            text_number = self.extract_text_from_box(page_num, scaled_rect_coords_number)
            if self.rect_coords_title:
                scaled_rect_coords_title = [coord / self.scale_factor for coord in self.rect_coords_title]
                text_title = self.extract_text_from_box(page_num, scaled_rect_coords_title)
            else:
                text_title = ""
            self.sheet_numbers_titles[index] = (page_num, text_number, text_title)
            self.entries_number[index].delete(0, tk.END)
            self.entries_number[index].insert(0, text_number)
            self.entries_title[index].delete(0, tk.END)
            self.entries_title[index].insert(0, text_title)

        self.draw_phase = None
        updated_count = len(self.active_retry_indices)
        for index in self.active_retry_indices:
            self.check_vars[index].set(0)
        if hasattr(self, "selection_window") and self.selection_window.winfo_exists():
            self.selection_window.deiconify()
            self.selection_window.lift()
            self.selection_window.focus_force()
        self.active_retry_indices = []
        self.set_status(f"Retry applied to {updated_count} checked rows. Review and save.")
        self.show_notice_popup("Retry Complete", "Checked rows have been reprocessed with your new boxes.")

    def show_notice_popup(self, title, message):
        popup = self._create_styled_popup(title, geometry="420x180")
        content_frame = self._popup_content_frame(popup)
        Label(content_frame, text=title, font=("Segoe UI", 12, "bold"), bg="#f8fafc", fg="#0f172a").pack(anchor="w", pady=(0, 8))
        Label(content_frame, text=message, bg="#f8fafc", fg="#334155", justify=tk.LEFT, wraplength=370).pack(anchor="w", pady=(0, 14))
        button_frame = Frame(content_frame, bg="#f8fafc")
        button_frame.pack(anchor="e")
        self._styled_button(button_frame, "OK", popup.destroy, primary=True).pack(side=tk.LEFT)

    def check_all(self):
        for var in self.check_vars:
            var.set(1)

    def uncheck_all(self):
        for var in self.check_vars:
            var.set(0)

    def check_drawings_by_letter(self, letter):
        for i, entry_number in enumerate(self.entries_number):
            sheet_number = entry_number.get().strip()
            if sheet_number and sheet_number.upper().startswith(letter):
                self.check_vars[i].set(1)

    def save_sheets(self, with_title):
        if not any(var.get() for var in self.check_vars):
            self.show_notice_popup("Warning", "No sheets selected to save.")
            return

        output_dir = filedialog.askdirectory()
        if not output_dir:
            return

        for index, var in enumerate(self.check_vars):
            if var.get() == 1:
                page_num, _, _ = self.sheet_numbers_titles[index]
                text_number = self.entries_number[index].get()
                text_title = self.entries_title[index].get()
                sanitized_text_number = self.sanitize_filename(text_number)
                if with_title:
                    sanitized_text_title = self.sanitize_filename(text_title)
                    output_path = os.path.join(output_dir, f"{sanitized_text_number} {sanitized_text_title}.pdf")
                else:
                    output_path = os.path.join(output_dir, f"{sanitized_text_number}.pdf")
                self.save_page_as_pdf(page_num, output_path)


    def preview_checked_pages(self):
        checked = [i for i, var in enumerate(self.check_vars) if var.get() == 1]
        if not checked:
            self.show_notice_popup("No Rows Selected", "Check at least one row to preview.")
            return
        if self.extraction_mode == "specs":
            first_page = self.spec_group_items[checked[0]][1][0]
        else:
            first_page = self.sheet_numbers_titles[checked[0]][0]
        self.page_number = first_page
        self.display_page(first_page - 1)
        self.set_status(f"Previewing first checked page: {first_page}")

    def save_specs_grouped(self):
        selected = [i for i, var in enumerate(self.check_vars) if var.get() == 1]
        if not selected:
            self.show_notice_popup("Warning", "No specs selected to save.")
            return
        output_dir = filedialog.askdirectory()
        if not output_dir:
            return

        groups = {}
        for index in selected:
            _, group_data = self.spec_group_items[index]
            pages = group_data["pages"]
            spec_number = self.spec_entries_number[index].get().strip()
            spec_name = self.spec_entries_title[index].get().strip()
            if not spec_number and not spec_name:
                continue
            groups[(spec_number, spec_name)] = pages

        if not groups:
            self.show_notice_popup("Nothing to Save", "No valid spec number/name pairs were selected.")
            return

        saved_count = 0
        failed_files = []
        for (spec_number, spec_name), pages in groups.items():
            safe_number = self.sanitize_filename(spec_number or "UNKNOWN_SPEC")
            safe_name = self.sanitize_filename(spec_name or "UNTITLED")
            output_path = self.build_safe_output_path(output_dir, safe_number, safe_name)
            try:
                new_doc = fitz.open()
                for page in sorted(set(pages)):
                    new_doc.insert_pdf(self.pdf_document, from_page=page - 1, to_page=page - 1)
                new_doc.save(output_path)
                saved_count += 1
            except Exception as exc:
                failed_files.append(f"{os.path.basename(output_path)} ({exc})")

        if failed_files:
            detail = "\n".join(failed_files[:4])
            if len(failed_files) > 4:
                detail += f"\n...and {len(failed_files) - 4} more."
            self.show_notice_popup(
                "Specs Saved with Warnings",
                f"Saved {saved_count} grouped spec PDF(s).\nFailed to save {len(failed_files)} file(s):\n{detail}",
            )
        else:
            self.show_notice_popup("Specs Saved", f"Saved {saved_count} grouped spec PDF(s).")

    def sanitize_filename(self, filename):
        sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename).strip()
        return re.sub(r'\s+', '_', sanitized)

    def build_safe_output_path(self, output_dir, safe_number, safe_name):
        base_name = f"{safe_number} - {safe_name}".strip(" -")
        if not base_name:
            base_name = "UNTITLED_SPEC"

        max_name_length = 180
        if len(base_name) > max_name_length:
            base_name = base_name[:max_name_length].rstrip(" ._-")

        output_path = os.path.join(output_dir, f"{base_name}.pdf")
        max_path_length = 245
        if len(output_path) > max_path_length:
            overflow = len(output_path) - max_path_length
            trimmed_name_len = max(40, len(base_name) - overflow - 1)
            base_name = base_name[:trimmed_name_len].rstrip(" ._-")
            output_path = os.path.join(output_dir, f"{base_name}.pdf")

        return output_path

    def save_page_as_pdf(self, page_number, output_path):
        new_doc = fitz.open()
        new_doc.insert_pdf(self.pdf_document, from_page=page_number - 1, to_page=page_number - 1)
        new_doc.save(output_path)

    def _on_mousewheel(self, event):
        if event.state & 0x0001:
            self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_shift_mousewheel(self, event):
        self.canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_checklist_mousewheel(self, event):
        if hasattr(self, "checklist_canvas"):
            if event.state & 0x0001:
                self.checklist_canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
            else:
                self.checklist_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")



def main():
    root = tk.Tk()
    PDFExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
