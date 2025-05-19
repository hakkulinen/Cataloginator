import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import asyncio
import os
import threading
from pathlib import Path
import aiohttp
import pandas as pd
from bs4 import BeautifulSoup
import urllib.parse
import re
import logging
import validators
from queue import Queue, Empty

# Logging Config
logging.basicConfig(
    filename="download_errors.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

async def fetch_html(session, url, retries=3, backoff_factor=1):
    """Fetch HTML content from a URL with retries."""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    for attempt in range(retries):
        try:
            async with session.get(url, headers=headers, timeout=30) as response:
                if response.status == 200:
                    return await response.text()
                elif response.status == 429:
                    wait_time = backoff_factor * (2 ** attempt)
                    logging.warning(f"Rate limit hit for {url}, waiting {wait_time}s")
                    await asyncio.sleep(wait_time)
                else:
                    logging.error(f"Failed to fetch {url}: Status {response.status}")
                    return None
        except Exception as e:
            logging.error(f"Error fetching {url} (attempt {attempt + 1}/{retries}): {e}")
            if attempt < retries - 1:
                wait_time = backoff_factor * (2 ** attempt)
                await asyncio.sleep(wait_time)
    logging.error(f"Failed to fetch {url} after {retries} attempts")
    return None

async def get_image_url(html_content, base_url):
    """Parse HTML to find the first .jpg image URL."""
    if not html_content:
        return None
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        img_tag = soup.find('img', src=re.compile(r'.*\.jpg$', re.I))
        if img_tag and img_tag['src']:
            img_url = urllib.parse.urljoin(base_url, img_tag['src'])
            return img_url
        logging.warning(f"No .jpg image found in {base_url}")
        return None
    except Exception as e:
        logging.error(f"Error parsing HTML for {base_url}: {e}")
        return None

async def download_image(session, img_url, filename, save_folder):
    """Download and save an image."""
    try:
        async with session.get(img_url, timeout=30) as response:
            if response.status == 200:
                safe_filename = re.sub(r'[^\w\-_\. ]', '_', filename)
                if not safe_filename.lower().endswith('.jpg'):
                    safe_filename += '.jpg'

                save_path = Path(save_folder) / safe_filename
                with open(save_path, 'wb') as f:
                    f.write(await response.read())
                print(f"Saved image: {save_path}")
            else:
                logging.warning(f"Failed to download {img_url}: Status {response.status}")
    except Exception as e:
        logging.error(f"Error downloading {img_url}: {e}")

async def process_row(session, row, save_folder, semaphore, progress_queue, row_index, total_rows):
    """Process a single row with semaphore and report progress."""
    async with semaphore:
        url = row[0]
        filename = str(row[1]).strip()

        if not url or not filename:
            logging.warning(f"Skipping row with empty URL or filename: {url}, {filename}")
            progress_queue.put((row_index + 1, total_rows))
            return

        if not validators.url(url):
            logging.warning(f"Invalid URL: {url}")
            progress_queue.put((row_index + 1, total_rows))
            return

        html_content = await fetch_html(session, url)
        if html_content:
            img_url = await get_image_url(html_content, url)
            if img_url:
                await download_image(session, img_url, filename, save_folder)
        progress_queue.put((row_index + 1, total_rows))

async def process_batch(session, batch, save_folder, semaphore, progress_queue, start_index, total_rows):
    """Process a batch of rows."""
    tasks = [
        process_row(session, row, save_folder, semaphore, progress_queue, start_index + i, total_rows)
        for i, (_, row) in enumerate(batch.iterrows())
    ]
    await asyncio.gather(*tasks)

async def main(excel_file, save_folder, progress_queue, max_concurrent=50, batch_size=100):
    """Main function to process the Excel file and download images."""
    Path(save_folder).mkdir(parents=True, exist_ok=True)

    try:
        df = pd.read_excel(excel_file, header=None)
        total_rows = len(df)
        if total_rows == 0:
            logging.error("Excel file is empty")
            return False, "Excel file is empty"
    except Exception as e:
        logging.error(f"Error reading Excel file: {e}")
        return False, str(e)

    batches = [df[i:i + batch_size] for i in range(0, len(df), batch_size)]
    semaphore = asyncio.Semaphore(max_concurrent)

    try:
        async with aiohttp.ClientSession() as session:
            for i, batch in enumerate(batches):
                print(f"Processing batch {i + 1}/{len(batches)}")
                start_index = i * batch_size
                await process_batch(session, batch, save_folder, semaphore, progress_queue, start_index, total_rows)
                await asyncio.sleep(1)
        return True, None
    except Exception as e:
        logging.warning(f"Error during download: {e}")
        return False, str(e)

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
        self.footer_label = tk.Label(root, text="made by vP v0.1", font=("Arial", 8), fg="gray")
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
                main(excel_file, save_folder, self.progress_queue)
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

if __name__ == "__main__":
    start_gui()