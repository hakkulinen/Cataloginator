import asyncio
import logging
import urllib.parse
import re
from pathlib import Path
import validators
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from bs4 import BeautifulSoup


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

async def download_image(session, img_url, filename, save_folder, date_str):
    """Download and save an image, then add date in yellow text at bottom right."""
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

                # Add date to the image
                try:
                    # Parse and format the date
                    date_obj = pd.to_datetime(date_str, errors='coerce')
                    if pd.isna(date_obj):
                        logging.warning(f"Invalid date format for {filename}: {date_str}")
                        return
                    formatted_date = date_obj.strftime('%Y-%m-%d')

                    # Open the image with Pillow
                    with Image.open(save_path) as img:
                        draw = ImageDraw.Draw(img)
                        try:
                            # Use a standard font, fall back to default if unavailable
                            font = ImageFont.truetype("arial.ttf", 30)
                        except:
                            font = ImageFont.load_default()

                        # Get text size and image dimensions
                        text_bbox = draw.textbbox((0, 0), formatted_date, font=font)
                        text_width = text_bbox[2] - text_bbox[0]
                        text_height = text_bbox[3] - text_bbox[1]
                        img_width, img_height = img.size

                        # Calculate position for bottom-right corner (with padding)
                        padding = 10
                        text_x = img_width - text_width - padding
                        text_y = img_height - text_height - padding

                        # Draw yellow text
                        draw.text((text_x, text_y), formatted_date, fill=(255, 255, 0), font=font)

                        # Save the modified image
                        img.save(save_path, 'JPEG')
                        print(f"Added date to image: {save_path}")
                except Exception as e:
                    logging.error(f"Error adding date to {save_path}: {e}")
            else:
                logging.warning(f"Failed to download {img_url}: Status {response.status}")
    except Exception as e:
        logging.error(f"Error downloading {img_url}: {e}")

async def process_row(session, row, save_folder, semaphore, progress_queue, row_index, total_rows):
    """Process a single row with semaphore and report progress."""
    async with semaphore:
        url = row[0]
        filename = str(row[1]).strip()
        date_str = str(row[2]).strip() if len(row) > 2 else ''

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
                await download_image(session, img_url, filename, save_folder, date_str)
        progress_queue.put((row_index + 1, total_rows))

async def process_batch(session, batch, save_folder, semaphore, progress_queue, start_index, total_rows):
    """Process a batch of rows."""
    tasks = [
        process_row(session, row, save_folder, semaphore, progress_queue, start_index + i, total_rows)
        for i, (_, row) in enumerate(batch.iterrows())
    ]
    await asyncio.gather(*tasks)