import logging

# Logging Config
logging.basicConfig(
    filename="errors.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
# Debug flag
DEBUG = False  # Set to False to disable debug logs