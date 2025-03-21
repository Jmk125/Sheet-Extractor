import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Scrollbar, IntVar, Frame, Canvas, Label, Button, Entry, Checkbutton
from PIL import Image, ImageTk
import os
import re

class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Sheet Extractor")
        
        self.pdf_path = ""
        self.page_number = 1
        self.rect_coords_number = None
        self.rect_coords_title = None
        self.pdf_document = None
        self.scale_factor = 1.0  # Increased scale factor to improve text extraction

        # Frame for buttons
        self.button_frame = Frame(root)
        self.button_frame.pack(fill=tk.X)
        
        self.upload_button = Button(self.button_frame, text="Upload PDF", command=self.upload_pdf)
        self.upload_button.pack(side=tk.LEFT)
        
        self.next_button = Button(self.button_frame, text="Next Drawing", command=self.next_drawing)
        self.next_button.pack(side=tk.LEFT)
        
        # Frame for canvas with scrollbars
        self.canvas_frame = Frame(root)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.canvas = Canvas(self.canvas_frame, bg="white")
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.scrollbar_y = Scrollbar(self.canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.scrollbar_x = Scrollbar(root, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.canvas.config(yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set)
        
        self.canvas.bind("<Button-1>", self.start_draw)
        self.canvas.bind("<B1-Motion>", self.draw_rect)
        self.canvas.bind("<ButtonRelease-1>", self.end_draw)
        
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_shift_mousewheel)

        # Ensure scrollbars are always visible and span the entire canvas
        self.root.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox(tk.ALL))
        self.scrollbar_y.lift(self.canvas)
        self.scrollbar_x.lift(self.canvas)

    def upload_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.pdf_path:
            self.page_number = 1
            self.load_first_page()
            # Show popup window with instruction
            self.show_instruction_popup()

    def next_drawing(self):
        if self.pdf_document and self.page_number < len(self.pdf_document):
            self.page_number += 1
            self.display_page(self.page_number - 1)

    def show_instruction_popup(self):
        instruction_popup = Toplevel(self.root)
        instruction_popup.title("Instruction")
        instruction_popup.geometry("300x100")
        
        label = Label(instruction_popup, text="Draw a box around the sheet number")
        label.pack(pady=20)
        
        ok_button = Button(instruction_popup, text="Ok", command=instruction_popup.destroy)
        ok_button.pack()

    def load_first_page(self):
        self.pdf_document = fitz.open(self.pdf_path)
        self.display_page(self.page_number - 1)

    def display_page(self, page_index):
        page = self.pdf_document.load_page(page_index)
        
        # Scale the page to the desired size
        zoom = self.scale_factor
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
        self.img_tk = ImageTk.PhotoImage(img)
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.img_tk)
        
        self.canvas.config(scrollregion=self.canvas.bbox(tk.ALL))
        self.canvas.config(width=self.canvas_frame.winfo_width(), height=self.canvas_frame.winfo_height())

        # Ensure scrollbars remain visible and span the entire canvas
        self.scrollbar_y.lift(self.canvas)
        self.scrollbar_x.lift(self.canvas)

        # Debug information about the page
        print(f"Page size: {page.rect}")

    def start_draw(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red")

    def draw_rect(self, event):
        current_x = self.canvas.canvasx(event.x)
        current_y = self.canvas.canvasy(event.y)
        self.canvas.coords(self.rect, self.start_x, self.start_y, current_x, current_y)

    def end_draw(self, event):
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        if self.rect_coords_number is None:
            # First box drawn around the sheet number
            self.rect_coords_number = (self.start_x, self.start_y, end_x, end_y)
            print(f"Selection box coordinates for number (scaled): {self.rect_coords_number}")

            # Extract text from the selected area
            scaled_rect_coords = [coord / self.scale_factor for coord in self.rect_coords_number]
            extracted_text_number = self.extract_text_from_box(self.page_number, scaled_rect_coords)
            if extracted_text_number:
                # Show confirmation popup for sheet number
                self.show_confirmation_popup_number(extracted_text_number)
            else:
                messagebox.showwarning("No Text", "No text found in the selected area.")
        else:
            # Second box drawn around the sheet title
            self.rect_coords_title = (self.start_x, self.start_y, end_x, end_y)
            print(f"Selection box coordinates for title (scaled): {self.rect_coords_title}")

            # Extract text from the selected area
            scaled_rect_coords = [coord / self.scale_factor for coord in self.rect_coords_title]
            extracted_text_title = self.extract_text_from_box(self.page_number, scaled_rect_coords)
            
            if extracted_text_title:
                # Consolidate text into a single line
                extracted_text_title = " ".join(extracted_text_title.split())
                # Show confirmation popup for sheet title
                self.show_confirmation_popup_title(extracted_text_title)
            else:
                messagebox.showwarning("No Text", "No text found in the selected area.")

    def extract_text_from_box(self, page_number, box):
        page = self.pdf_document.load_page(page_number - 1)
        rect = fitz.Rect(box)
        print(f"Extracting text from rect: {rect}")  # Debug information
        
        text = page.get_text("text", clip=rect).strip()
        
        print(f"Extracted text: {text}")  # Debug information
        
        return text

    def show_confirmation_popup_number(self, extracted_text_number):
        confirmation_popup = Toplevel(self.root)
        confirmation_popup.title("Confirm Sheet Number")
        confirmation_popup.geometry("300x200")
        
        label = Label(confirmation_popup, text=f"Extracted text: {extracted_text_number}\nIs this correct?")
        label.pack(pady=10)
        
        button_frame = Frame(confirmation_popup)
        button_frame.pack(pady=10)
        
        cancel_button = Button(button_frame, text="Cancel", command=confirmation_popup.destroy)
        cancel_button.pack(side=tk.LEFT, padx=10)
        
        confirm_button = Button(button_frame, text="Confirm & Draw Title Box", command=lambda: [self.activate_title_box_drawing(), confirmation_popup.destroy()])
        confirm_button.pack(side=tk.LEFT)
        
        skip_button = Button(button_frame, text="Skip Title", command=lambda: [self.skip_title_drawing(), confirmation_popup.destroy()])
        skip_button.pack(side=tk.LEFT)

    def skip_title_drawing(self):
        self.rect_coords_number_confirmed = True
        self.rect_coords_title = None
        self.process_all_pages()

    def activate_title_box_drawing(self):
        self.rect_coords_number_confirmed = True

    def show_confirmation_popup_title(self, extracted_text_title):
        confirmation_popup = Toplevel(self.root)
        confirmation_popup.title("Confirm Sheet Title")
        confirmation_popup.geometry("300x150")
        
        self.extracted_text_title = extracted_text_title  # Store the extracted title text
        
        label = Label(confirmation_popup, text=f"Extracted text: {extracted_text_title}\nIs this correct?")
        label.pack(pady=20)
        
        button_frame = Frame(confirmation_popup)
        button_frame.pack(pady=10)
        
        cancel_button = Button(button_frame, text="Cancel", command=confirmation_popup.destroy)
        cancel_button.pack(side=tk.LEFT, padx=10)
        
        confirm_button = Button(button_frame, text="Confirm & Extract Sheets", command=lambda: [self.process_all_pages(), confirmation_popup.destroy()])
        confirm_button.pack(side=tk.LEFT)

    def process_all_pages(self):
        self.sheet_numbers_titles = []
        for page_num in range(1, self.pdf_document.page_count + 1):
            scaled_rect_coords_number = [coord / self.scale_factor for coord in self.rect_coords_number]
            text_number = self.extract_text_from_box(page_num, scaled_rect_coords_number)
            if self.rect_coords_title:
                scaled_rect_coords_title = [coord / self.scale_factor for coord in self.rect_coords_title]
                text_title = self.extract_text_from_box(page_num, scaled_rect_coords_title)
            else:
                text_title = ""
            self.sheet_numbers_titles.append((page_num, text_number, text_title))
        
        self.show_sheet_selection()

    def show_sheet_selection(self):
        selection_window = Toplevel(self.root)
        selection_window.title("Select Sheets to Extract")
        selection_window.geometry("600x600")  # Set larger initial size
        
        # Frame for buttons at the top
        button_frame_top = Frame(selection_window)
        button_frame_top.pack(side=tk.TOP, pady=10, fill=tk.X)
        
        check_all_button = Button(button_frame_top, text="Check All", command=self.check_all)
        check_all_button.pack(side=tk.LEFT, padx=5)
        
        uncheck_all_button = Button(button_frame_top, text="Uncheck All", command=self.uncheck_all)
        uncheck_all_button.pack(side=tk.LEFT, padx=5)
        
        a_drawings_button = Button(button_frame_top, text="A Drawings", command=lambda: self.check_drawings_by_letter('A'))
        a_drawings_button.pack(side=tk.LEFT, padx=5)
        
        c_drawings_button = Button(button_frame_top, text="C Drawings", command=lambda: self.check_drawings_by_letter('C'))
        c_drawings_button.pack(side=tk.LEFT, padx=5)
        
        s_drawings_button = Button(button_frame_top, text="S Drawings", command=lambda: self.check_drawings_by_letter('S'))
        s_drawings_button.pack(side=tk.LEFT, padx=5)
        
        # Frame for the checklist with scrollbars
        checklist_frame = Frame(selection_window)
        checklist_frame.pack(fill=tk.BOTH, expand=True)
        
        self.checklist_canvas = Canvas(checklist_frame)
        self.checklist_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = Scrollbar(checklist_frame, orient=tk.VERTICAL, command=self.checklist_canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.checklist_canvas.config(yscrollcommand=scrollbar.set)
        
        checklist_inner_frame = Frame(self.checklist_canvas)
        self.checklist_canvas.create_window((0, 0), window=checklist_inner_frame, anchor=tk.NW)
        
        checklist_inner_frame.bind("<Configure>", lambda e: self.checklist_canvas.config(scrollregion=self.checklist_canvas.bbox(tk.ALL)))
        self.checklist_canvas.bind_all("<MouseWheel>", self._on_checklist_mousewheel)
        
        self.check_vars = []
        self.checkbuttons = []
        self.entries_number = []
        self.entries_title = []
        
        for page_num, text_number, text_title in self.sheet_numbers_titles:
            var = IntVar()
            frame = Frame(checklist_inner_frame)
            frame.pack(anchor=tk.W, fill=tk.X)
            
            checkbutton = Checkbutton(frame, variable=var)
            checkbutton.pack(side=tk.LEFT, padx=5)
            self.check_vars.append(var)
            self.checkbuttons.append(checkbutton)
            
            entry_number = Entry(frame)
            entry_number.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            entry_number.insert(0, text_number)
            self.entries_number.append(entry_number)
            
            entry_title = Entry(frame, width=40)  # Adjusted width to be twice as wide
            entry_title.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            entry_title.insert(0, text_title)
            self.entries_title.append(entry_title)
        
        # Frame for buttons at the bottom
        button_frame_bottom = Frame(selection_window)
        button_frame_bottom.pack(side=tk.BOTTOM, pady=10)
        
        save_number_button = Button(button_frame_bottom, text="Save Sheets with Number", width=25, command=lambda: self.save_sheets(with_title=False))
        save_number_button.pack(side=tk.LEFT, padx=5)
        
        save_number_title_button = Button(button_frame_bottom, text="Save Sheets with Number and Title", width=30, command=lambda: self.save_sheets(with_title=True))
        save_number_title_button.pack(side=tk.LEFT, padx=5)

    def check_all(self):
        for var in self.check_vars:
            var.set(1)

    def uncheck_all(self):
        for var in self.check_vars:
            var.set(0)

    def check_drawings_by_letter(self, letter):
        for var, checkbutton in zip(self.check_vars, self.checkbuttons):
            if checkbutton.cget("text").startswith(letter):
                var.set(1)

    def save_sheets(self, with_title):
        if not any(var.get() for var in self.check_vars):
            messagebox.showwarning("Warning", "No sheets selected to save.")
            return
        
        output_dir = filedialog.askdirectory()
        if not output_dir:
            return
        
        for index, var in enumerate(self.check_vars):
            if var.get() == 1:
                page_num, text_number, text_title = self.sheet_numbers_titles[index]
                text_number = self.entries_number[index].get()  # Get the potentially modified sheet number
                text_title = self.entries_title[index].get()  # Get the potentially modified sheet title
                sanitized_text_number = self.sanitize_filename(text_number)
                if with_title:
                    sanitized_text_title = self.sanitize_filename(text_title)
                    output_path = os.path.join(output_dir, f"{sanitized_text_number} {sanitized_text_title}.pdf")
                else:
                    output_path = os.path.join(output_dir, f"{sanitized_text_number}.pdf")
                self.save_page_as_pdf(page_num, output_path)
                print(f"Saved page {page_num} as {output_path}")

    def sanitize_filename(self, filename):
        # Replace invalid characters with underscores and trim leading/trailing whitespace
        sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename).strip()
        # Replace spaces with underscores
        return re.sub(r'\s+', '_', sanitized)

    def save_page_as_pdf(self, page_number, output_path):
        new_doc = fitz.open()  # Create a new PDF
        new_doc.insert_pdf(self.pdf_document, from_page=page_number - 1, to_page=page_number - 1)
        new_doc.save(output_path)
    
    def _on_mousewheel(self, event):
        if event.state & 0x0001:  # Shift key is held
            self.canvas.xview_scroll(int(-1*(event.delta/120)), "units")
        else:
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_shift_mousewheel(self, event):
        self.canvas.xview_scroll(int(-1*(event.delta/120)), "units")

    def _on_checklist_mousewheel(self, event):
        if event.state & 0x0001:  # Shift key is held
            self.checklist_canvas.xview_scroll(int(-1*(event.delta/120)), "units")
        else:
            self.checklist_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

def main():
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()