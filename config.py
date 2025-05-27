import logging

# Logging Config
logging.basicConfig(
    filename="errors.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def flip(flipper):
    flipper = (flipper + 1) % 2  # Toggles between 0 and 1
    return flipper

# Debug flag
DEBUG = False  # Set to False to disable debug logs