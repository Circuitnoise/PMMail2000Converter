"""
PMMail2000Converter
-------------------

A converter tool that transforms legacy PMMail 2000 archives into EML files,
reconstructing folder names from ACCT.INI and FOLDER.INI files.

CLI command: `pmmail-convert`
"""

import logging

__version__ = "1.0.0"
__author__ = "Your Name"
__email__ = "your.email@example.com"

# Set up a default NullHandler to avoid logging errors if not configured by the user
logging.getLogger(__name__).addHandler(logging.NullHandler())

from .msg2eml import main

__all__ = ["main"]
