# log.py

import os
import sys
import logging
from logging.handlers import RotatingFileHandler

# === Directory setup ===
LOG_DIR = os.path.join(os.path.dirname(__file__), "logs")
os.makedirs(LOG_DIR, exist_ok=True)

# === Logger configuration ===
logger = logging.getLogger("WiFITestSuite")
logger.setLevel(logging.DEBUG)
logger.propagate = False  # don’t bubble up to root

# Internal flag to ensure handlers are only added once
_handlers_initialized = False

def delete_old_logs():
    """Remove any .log files in LOG_DIR before starting fresh."""
    for fname in os.listdir(LOG_DIR):
        if fname.endswith(".log"):
            try:
                os.remove(os.path.join(LOG_DIR, fname))
            except OSError:
                # If deletion fails, we’ll catch it downstream
                pass

def setup_log_handlers():
    """
    Install file and console handlers exactly once.
    Returns True on success, False on failure.
    """
    global _handlers_initialized
    if _handlers_initialized:
        return True

    try:
        # Rotating handlers, append mode, UTF-8
        handlers = [
            RotatingFileHandler(
                os.path.join(LOG_DIR, "runtime.log"),
                maxBytes=10 * 1024 * 1024,
                backupCount=5,
                encoding="utf-8"
            ),
            RotatingFileHandler(
                os.path.join(LOG_DIR, "errors.log"),
                maxBytes=5 * 1024 * 1024,
                backupCount=3,
                encoding="utf-8"
            ),
            RotatingFileHandler(
                os.path.join(LOG_DIR, "crash.log"),
                maxBytes=2 * 1024 * 1024,
                backupCount=2,
                encoding="utf-8"
            ),
            logging.StreamHandler(sys.stderr),  # console fallback
        ]

        # Levels: INFO+ → runtime, WARNING+ → errors, ERROR+ → crash, DEBUG+ → console
        handlers[0].setLevel(logging.INFO)
        handlers[1].setLevel(logging.WARNING)
        handlers[2].setLevel(logging.ERROR)
        handlers[3].setLevel(logging.DEBUG)

        fmt = logging.Formatter(
            "%(asctime)s %(name)s [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S"
        )
        for h in handlers:
            h.setFormatter(fmt)
            logger.addHandler(h)

        _handlers_initialized = True
        return True

    except OSError:
        # If file handlers failed, still add console handler for errors
        if not any(isinstance(h, logging.StreamHandler) for h in logger.handlers):
            console = logging.StreamHandler(sys.stderr)
            console.setLevel(logging.DEBUG)
            console.setFormatter(logging.Formatter("%(name)s [%(levelname)s] %(message)s"))
            logger.addHandler(console)
        logger.error("Failed to initialize file handlers for logging", exc_info=True)
        return False

# === Uncaught exception hook ===
def _handle_exception(exc_type, exc_value, exc_tb):
    if issubclass(exc_type, KeyboardInterrupt):
        # Delegate to default hook for Ctrl+C
        sys.__excepthook__(exc_type, exc_value, exc_tb)
    else:
        logger.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_tb))

sys.excepthook = _handle_exception

# === Initialization ===
def initialize_logging(clean_logs: bool = True) -> bool:
    """
    Call this once at program start.
    
    Args:
        clean_logs: if True, deletes old .log files before setup.
    
    Returns:
        True if handlers are in place; False otherwise.
    """
    if clean_logs:
        delete_old_logs()
    return setup_log_handlers()
