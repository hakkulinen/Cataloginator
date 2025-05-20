import asyncio
from pathlib import Path
import aiohttp
import pandas as pd
import logging

from webdownloader import process_batch
import gui

# Logging Config
logging.basicConfig(
    filename="download_errors.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

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

if __name__ == "__main__":
    gui.start_gui()