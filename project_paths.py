"""
Hardcoded project paths for local runs and Google Colab.

Colab usage:
    exec(open("/content/Siemens-road/run_pipeline.py").read())

The Colab root is hardcoded below. When that folder does not exist (e.g. on
Windows), the directory containing this file is used instead.
"""

from __future__ import annotations

from pathlib import Path

# Hardcoded Colab root — upload the project to /content/Siemens-road/
_COLAB_BASE_DIR = Path("/content/Siemens-Road")
_SCRIPT_DIR = Path(__file__).resolve().parent

BASE_DIR = next((path for path in _COLAB_BASE_DIRS if path.is_dir()), _SCRIPT_DIR)

INPUT_DIR = Path("/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT Siemens/Siemens Road/input")
PROCESSING_DIR = Path("/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT Siemens/Siemens Road/processing")
OUTPUT_DIR = Path("/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT Siemens/Siemens Road/output")


def ensure_project_dirs() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


