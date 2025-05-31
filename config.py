import logging

# Logging Config
logging.basicConfig(
    filename="errors.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
#int value for the logic of printing the defects in the top right of the image. Might change with new defects being added
MAX_DEFECT_LENGTH = 17

def flip(flipper):
    flipper = (flipper + 1) % 2  # Toggles between 0 and 1
    return flipper

# Debug flag
DEBUG = False  # Set to False to disable debug logs