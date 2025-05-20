import asyncio
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import logging
from queue import Queue, Empty

from webdownloader import async_download_manager
import config


class ImageDownloaderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Pictures Extractor")
        self.root.geometry("400x400")
        self.root.resizable(False, False)

        # Excel file selection
        self.label_excel = tk.Label(root, text="Excel File Destination:")
        self.label_excel.pack(pady=10)
        self.entry_excel = tk.Entry(root, width=40)
        self.entry_excel.pack()
        self.button_browse_excel = tk.Button(root, text="Browse", command=self.browse_excel)
        self.button_browse_excel.pack(pady=10)
        self.button_browse_excel.pack()

        # Save folder selection
        self.label_folder = tk.Label(root, text="Save Folder Destination:")
        self.label_folder.pack(pady=10)
        self.entry_folder = tk.Entry(root, width=40)
        self.entry_folder.pack()
        self.button_browse_folder = tk.Button(root, text="Browse", command=self.browse_folder)
        self.button_browse_folder.pack(pady=10)
        self.button_browse_folder.pack()

        # Download button
        self.button_download = tk.Button(root, text="Download", command=self.start_download)
        self.button_download.pack(pady=20)

        # Progress bar
        self.progress_bar = ttk.Progressbar(root, orient="horizontal", length=330, mode="determinate")
        self.progress_bar.pack(pady=10)
        self.progress_label = tk.Label(root, text="")
        self.progress_label.pack()

        # Footer label
        self.footer_label = tk.Label(root, text="made by vP v0.2", font=("Arial", 8), fg="gray")
        self.footer_label.pack(side="bottom", pady=5)

        # Queue for progress updates
        self.progress_queue = Queue()

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

def start_gui():
    root = tk.Tk()
    app = ImageDownloaderGUI(root)
    root.mainloop()