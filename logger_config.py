import logging
import os

# Mapping log levels to icons
LEVEL_ICONS = {
    logging.DEBUG: "🐞",
    logging.INFO: "ℹ️",
    logging.WARNING: "⚠️",
    logging.ERROR: "❌",
    logging.CRITICAL: "💥"
}

class IconFormatter(logging.Formatter):
    """Custom formatter that adds icons based on log level."""
    def format(self, record):
        icon = LEVEL_ICONS.get(record.levelno, "")
        record.msg = f"{icon} {record.msg}"
        return super().format(record)

def get_logger(name: str):
    """
    Returns a logger instance with UTF-8 + icon support.
    """
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    if not logger.handlers:  # Prevent adding handlers multiple times
        formatter = IconFormatter(
            fmt="%(asctime)s | %(levelname)-8s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )

        # Ensure logs directory exists
        os.makedirs("logs", exist_ok=True)

        # File handler (UTF-8) writing to logs/log.txt
        file_handler = logging.FileHandler(
            os.path.join("logs", "log.txt"),
            mode='a',
            encoding='utf-8',
            errors='replace'
        )
        file_handler.setFormatter(formatter)

        # Console handler (optional)
        # console_handler = logging.StreamHandler()
        # console_handler.setFormatter(formatter)

        logger.addHandler(file_handler)
        # logger.addHandler(console_handler)

    return logger