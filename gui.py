import asyncio
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import logging
from queue import Queue, Empty
from PIL import Image, ImageTk
import shutil
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter

from webdownloader import async_download_manager
import config

class ImageDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Pictures Extractor")
        self.root.geometry("500x500")
        self.root.resizable(False, False)

        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(pady=10, fill="both", expand=True)

        # Download tab
        self.download_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.download_frame, text="Download")
        self.setup_download_tab()

        # Catalog tab
        self.catalog_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.catalog_frame, text="Catalog")
        self.setup_catalog_tab()

        # Footer label
        self.footer_label = tk.Label(self.root, text="made by vP v0.3", font=("Arial", 8), fg="gray")
        self.footer_label.pack(side="bottom", pady=5)

        # Queue for progress updates
        self.progress_queue = Queue()

    def setup_download_tab(self):
        # Excel file selection
        self.label_excel = tk.Label(self.download_frame, text="Excel File Destination:")
        self.label_excel.pack(pady=10)
        self.entry_excel = tk.Entry(self.download_frame, width=40)
        self.entry_excel.pack()
        self.button_browse_excel = tk.Button(self.download_frame, text="Browse", command=self.browse_excel)
        self.button_browse_excel.pack(pady=10)

        # Save folder selection
        self.label_folder = tk.Label(self.download_frame, text="Save Folder Destination:")
        self.label_folder.pack(pady=10)
        self.entry_folder = tk.Entry(self.download_frame, width=40)
        self.entry_folder.pack()
        self.button_browse_folder = tk.Button(self.download_frame, text="Browse", command=self.browse_folder)
        self.button_browse_folder.pack(pady=10)

        # Download button
        self.button_download = tk.Button(self.download_frame, text="Download", command=self.start_download)
        self.button_download.pack(pady=20)

        # Progress bar
        self.progress_bar = ttk.Progressbar(self.download_frame, orient="horizontal", length=330, mode="determinate")
        self.progress_bar.pack(pady=10)
        self.progress_label = tk.Label(self.download_frame, text="")
        self.progress_label.pack()

    def setup_catalog_tab(self):
        # Save folder selection
        self.label_catalog_folder = tk.Label(self.catalog_frame, text="Catalog Folder Destination:")
        self.label_catalog_folder.pack(pady=10)
        self.entry_catalog_folder = tk.Entry(self.catalog_frame, width=40)
        self.entry_catalog_folder.pack()
        self.button_browse_catalog_folder = tk.Button(self.catalog_frame, text="Browse", command=self.browse_catalog_folder)
        self.button_browse_catalog_folder.pack(pady=10)

        # Begin button
        self.button_begin = tk.Button(self.catalog_frame, text="Begin", command=self.start_cataloging)
        self.button_begin.pack(pady=20, side="bottom")

    def browse_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.entry_excel.delete(0, tk.END)
            self.entry_excel.insert(0, file_path)

    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.entry_folder.delete(0, tk.END)
            self.entry_folder.insert(0, folder_path)

    def browse_catalog_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.entry_catalog_folder.delete(0, tk.END)
            self.entry_catalog_folder.insert(0, folder_path)

    def start_download(self):
        excel_file = self.entry_excel.get()
        save_folder = self.entry_folder.get()

        if not excel_file or not save_folder:
            messagebox.showerror("Error", "Please select both Excel file and save folder.")
            return

        # Disable button and reset progress
        self.button_download.config(state="disabled")
        self.progress_bar["value"] = 0
        self.progress_label.config(text="Starting download...")

        # Start download in a separate thread
        threading.Thread(
            target=self.run_download, args=(excel_file, save_folder), daemon=True
        ).start()
        # Start polling for progress
        self.update_progress()

    def run_download(self, excel_file, save_folder):
        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            success, error = loop.run_until_complete(
                async_download_manager(excel_file, save_folder, self.progress_queue)
            )
            loop.close()
            self.root.after(0, self.show_result, success, error, save_folder)
        except Exception as e:
            self.root.after(0, self.show_result, False, str(e), save_folder)

    def update_progress(self):
        try:
            while True:
                current, total = self.progress_queue.get_nowait()
                percentage = (current / total) * 100
                self.progress_bar["value"] = percentage
                self.progress_label.config(text=f"Processed {current}/{total} rows")
        except Empty:
            pass
        if self.button_download["state"] == "disabled":
            self.root.after(100, self.update_progress)

    def show_result(self, success, error, save_folder):
        self.button_download.config(state="normal")
        self.progress_label.config(text="")
        self.progress_bar["value"] = 0

        if success:
            image_count = len(
                [f for f in os.listdir(save_folder) if os.path.isfile(os.path.join(save_folder, f))]
            )
            messagebox.showinfo(
                "Success", f"Download completed successfully!\n{image_count} images downloaded."
            )
            logging.debug("Script finished")
        else:
            messagebox.showerror("Error", f"An error occurred: {error}")

    def start_cataloging(self):
        catalog_folder = self.entry_catalog_folder.get()
        if not catalog_folder:
            messagebox.showerror("Error", "Please select a catalog folder.")
            return

        # Create processed and hold folders
        processed_folder = Path("./") / "processed"
        hold_folder = Path("./") / "hold"
        processed_folder.mkdir(exist_ok=True)
        hold_folder.mkdir(exist_ok=True)

        # Get list of images
        image_extensions = ('.jpg', '.jpeg', '.png')
        images = [
            f for f in os.listdir(catalog_folder)
            if os.path.isfile(os.path.join(catalog_folder, f)) and f.lower().endswith(image_extensions)
        ]

        if not images:
            messagebox.showinfo("Info", "No images found in the selected folder.")
            return

        # Initialize Excel report
        self.initialize_excel_report(catalog_folder)

        # Open cataloging window
        self.open_cataloging_window(catalog_folder, images, processed_folder, hold_folder)

    def initialize_excel_report(self, catalog_folder):
        self.excel_path = Path("./") / "catalog_report.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Catalog Report"
        headers = [
            "bwu", "region", "Outlet number", "Scene id", "BWU type",
            "Legislation issue", "Switched off", "Header", "Screen/SAS", "Exterior frame",
            "Shelves light", "Number", "Shelfstrip (base)", "Number", "Shelfstrip (insert)",
            "Number", "Adjust height of shelves", "No holders of packs", "Other"
        ]
        for col, header in enumerate(headers, 1):
            ws[f"{get_column_letter(col)}1"] = header
        wb.save(self.excel_path)
        logging.debug(f"Initialized Excel report at {self.excel_path}")

    def open_cataloging_window(self, catalog_folder, images, processed_folder, hold_folder):
        catalog_window = tk.Toplevel(self.root)
        catalog_window.title("Catalog Images")
        try:
            catalog_window.state('zoomed')  # Try standard maximize method
        except tk.TclError:
            # Fallback: Set window size to screen dimensions
            screen_width = catalog_window.winfo_screenwidth()
            screen_height = catalog_window.winfo_screenheight()
            catalog_window.geometry(f"{screen_width}x{screen_height}+0+0")
        catalog_window.resizable(True, True)  # Allow resizing

        # Exit maximized window with Escape key
        catalog_window.bind('<Escape>', lambda e: catalog_window.destroy())

        self.current_image_index = 0

        # Main frame
        main_frame = ttk.Frame(catalog_window)
        main_frame.pack(fill="both", expand=True)

        # Image display on left
        self.image_label = tk.Label(main_frame)
        self.image_label.pack(side="left", padx=20, pady=20)

        # Right frame for buttons and checkboxes
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side="right", padx=20, pady=20, fill="y")

        # BWU Type (top right, mutually exclusive)
        bwu_frame = ttk.Frame(right_frame)
        bwu_frame.pack(side="top", anchor="center")
        bwu_label = tk.Label(bwu_frame, text="BWU Type:", font=("arial.ttf", 12))
        bwu_label.pack(side="left")
        self.bwu_var = tk.StringVar(value="")
        bwu_types = [
            "PRO", "X", "Mini", "X flap", "Mini flap", "A2", "Pr 12/15",
            "SS Flaps", "Door Slim 12/15", "Door Oval 12/15", "Other"
        ]
        for bwu_type in bwu_types:
            rb = tk.Radiobutton(bwu_frame, text=bwu_type, variable=self.bwu_var, value=bwu_type, font=("arial.ttf", 12), padx=2, pady=5)
            rb.pack(side="left", padx=2)

        # Detected Defects (middle right, non-mutually exclusive)
        defects_frame = ttk.Frame(right_frame)
        defects_frame.pack(pady=40, anchor="center")
        defects_label = tk.Label(defects_frame, text="Detected defects:", font=("arial.ttf", 15))
        defects_label.pack(anchor="w")
        self.defect_vars = {}
        self.number_entries = {}
        defect_types = [
            "Legislation issue", "Switched-off", "Header", "Screen/SAS", "Exterior frame",
            "Shelves light", "Shelfstrip (base)", "Shelfstrip (insert)", "Adjust height of shelves",
            "No holders for packs", "Other"
        ]
        for defect in defect_types:
            var = tk.BooleanVar(value=False)
            self.defect_vars[defect] = var
            row_frame = ttk.Frame(defects_frame)
            row_frame.pack(fill="x", pady=2)
            cb = tk.Checkbutton(row_frame, text=defect, variable=var, font=("arial.ttf", 12), padx=5, pady=5)
            cb.pack(side="left")
            if defect in ["Shelves light", "Shelfstrip (base)", "Shelfstrip (insert)"]:
                entry = tk.Entry(row_frame, width=7, font=("arial.ttf", 12))
                entry.pack(side="left", padx=5)
                self.number_entries[defect] = entry

        # Submit and Hold buttons (bottom right)
        button_frame = ttk.Frame(right_frame)
        button_frame.pack(side="bottom", anchor="se")
        submit_button = tk.Button(
            button_frame, text="Submit", bg="green", fg="white", width=14, font=("arial.ttf", 14),
            command=lambda: self.process_image(
                catalog_folder, images, processed_folder, hold_folder, catalog_window, "processed"
            )
        )
        submit_button.pack(side="bottom", pady=10)
        hold_button = tk.Button(
            button_frame, text="Hold", bg="red", fg="white", width=14, font=("arial.ttf", 14),
            command=lambda: self.process_image(
                catalog_folder, images, processed_folder, hold_folder, catalog_window, "hold"
            )
        )
        hold_button.pack(side="bottom", pady=10)

        # Load first image
        self.load_image(catalog_folder, images, catalog_window)

    def load_image(self, catalog_folder, images, catalog_window):
        if self.current_image_index >= len(images):
            # Clear image and show message
            self.image_label.config(image=None)
            self.image_label.config(text="No more images to catalog!", font=("arial.ttf", 24))
            return

        image_path = os.path.join(catalog_folder, images[self.current_image_index])
        try:
            # Load and resize image to fit window
            img = Image.open(image_path)
            max_size = (800, 600)  # Adjust based on screen size
            img.thumbnail(max_size, Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self.image_label.config(image=photo, text="")
            self.image_label.image = photo  # Keep reference
            # Reset checkboxes and number entries
            self.bwu_var.set("")
            for var in self.defect_vars.values():
                var.set(False)
            for entry in self.number_entries.values():
                entry.delete(0, tk.END)
        except Exception as e:
            logging.error(f"Error loading image {image_path}: {e}")
            self.image_label.config(text="Error loading image", font=("Arial", 20))

    def process_image(self, catalog_folder, images, processed_folder, hold_folder, catalog_window, action):
        if self.current_image_index >= len(images):
            return

        current_image = images[self.current_image_index]
        source_path = os.path.join(catalog_folder, current_image)
        dest_folder = processed_folder if action == "processed" else hold_folder
        dest_path = os.path.join(dest_folder, current_image)

        try:
            # Save to Excel only on Submit
            if action == "processed":
                self.save_to_excel(current_image)
            # Move image
            shutil.move(source_path, dest_path)
            logging.debug(f"Moved {current_image} to {action} folder")
        except Exception as e:
            logging.error(f"Error moving {current_image} to {action} folder: {e}")
            messagebox.showerror("Error", f"Failed to move {current_image}: {e}")
            return

        # Move to next image
        self.current_image_index += 1
        self.load_image(catalog_folder, images, catalog_window)

    def save_to_excel(self, image_name):
        try:
            wb = openpyxl.load_workbook(self.excel_path)
            ws = wb.active
            row = ws.max_row + 1

            # Parse filename (assuming format: bwu.region.outlet.scene.jpg)
            parts = image_name.rsplit('.', 4)  # Split on last 4 dots
            if len(parts) >= 4:
                bwu, region, outlet, scene = parts[:4]
            else:
                bwu = region = outlet = scene = ""

            # Gather data
            data = [
                bwu, region, outlet, scene, self.bwu_var.get()
            ]
            defect_types = [
                "Legislation issue", "Switched-off", "Header", "Screen/SAS", "Exterior frame",
                "Shelves light", "Shelfstrip (base)", "Shelfstrip (insert)", "Adjust height of shelves",
                "No holders for packs", "Other"
            ]
            for defect in defect_types:
                data.append("Y" if self.defect_vars[defect].get() else "")
                if defect in self.number_entries:
                    number = self.number_entries[defect].get().strip()
                    data.append(number if number else "")

            # Write to Excel
            for col, value in enumerate(data, 1):
                ws[f"{get_column_letter(col)}{row}"] = value

            wb.save(self.excel_path)
            logging.debug(f"Saved catalog data for {image_name} to Excel")
        except Exception as e:
            logging.error(f"Error saving to Excel for {image_name}: {e}")
            messagebox.showerror("Error", f"Failed to save Excel data: {e}")

def start_gui():
    root = tk.Tk()
    app = ImageDownloaderGUI(root)
    root.mainloop()