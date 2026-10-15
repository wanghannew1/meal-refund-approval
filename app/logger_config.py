"""
Logging configuration for the meal refund approval application.

Sets up UTF-8 logging to both file and console, with Chinese
character support for terminal output.
"""

import sys
import logging
from pathlib import Path

# Force UTF-8 on stdout/stderr so Chinese logs display correctly in any terminal
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

# Configure logging
LOG_DIR = Path(__file__).parent.parent / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "app.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8", mode="a"),
        logging.StreamHandler(sys.stdout),
    ],
)

logger = logging.getLogger("MealRefundApproval")
