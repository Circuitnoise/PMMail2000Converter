import os
import re
import shutil
import logging
from pathlib import Path
from tqdm import tqdm
import extract_msg
import struct

# --------------------------------------------------------
# CONFIGURATION
# Change these paths as needed
# --------------------------------------------------------

SOURCE_DIR = Path("/Volumes/blueprint Works/Blueprint Software Works/PMMail 2000")
TARGET_DIR = Path("/Volumes/Convert")
LOG_FILE = TARGET_DIR / "conversion_log.txt"

# --------------------------------------------------------
# INITIALIZE LOGGING
# --------------------------------------------------------
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# --------------------------------------------------------
# HELPER FUNCTIONS
# --------------------------------------------------------

def sanitize_name(name: str) -> str:
    """Removes invalid filesystem characters and trims whitespace."""
    return re.sub(r'[\\/*?:"<>|]', "_", name.strip())


def read_acct_name(act_dir: Path) -> str:
    """Reads and extracts the account name from ACCT.INI."""
    ini_file = act_dir / "ACCT.INI"
    if not ini_file.exists():
        return sanitize_name(act_dir.stem)

    try:
        raw = ini_file.read_bytes()
        # Replace all null bytes with delimiters
        text = raw.replace(b"\x00", b"|").decode("latin1", errors="ignore")

        # Look for ACCTNAME and its following block
        m = re.search(r"ACCTNAME\|+([^|]+)", text)
        if m:
            name = m.group(1).strip()
            if name:
                return sanitize_name(name)

        # Fallback: sometimes ACCTNAME appears without delimiters
        m = re.search(r"ACCTNAME\s*([A-Za-z0-9@._\-\s]+)", text)
        if m:
            return sanitize_name(m.group(1).strip())

    except Exception as e:
        logging.error(f"Error reading {ini_file}: {e}", exc_info=True)

    # Fallback to directory name
    return sanitize_name(act_dir.stem)


def read_folder_name(fld_dir: Path) -> str:
    """Extracts folder names from FOLDER.INI – robust against encoding issues."""
    ini_file = fld_dir / "FOLDER.INI"
    if not ini_file.exists():
        logging.warning(f"FOLDER.INI missing in {fld_dir}")
        return sanitize_name(fld_dir.stem)

    try:
        raw = ini_file.read_bytes()
        text = raw.replace(b"\x00", b"").decode(errors="ignore").strip()

        # Replace non-printable characters
        cleaned = re.sub(r"[^\x20-\x7E]", "|", text)

        # 1️⃣ Classic format ("Inboxfi1")
        match = re.search(r"^([\wÄÖÜäöüß\s\-]+?)(?:fi|\ufb01)\d", cleaned)
        if match:
            return sanitize_name(match.group(1))

        # 2️⃣ Normalized or broken encoding variants – everything until a number or semicolon
        match = re.search(r"^[!]?([A-Za-zÄÖÜäöüß\s\-]+?)(?=[0-9;|])", cleaned)
        if match:
            return sanitize_name(match.group(1))

        # 3️⃣ Fallback – first segment before a delimiter
        match = re.search(r"^[!]?([^|;\r\n]+)", cleaned)
        if match:
            return sanitize_name(match.group(1))

        logging.warning(f"No valid folder name found in {ini_file}")

    except Exception as e:
        logging.error(f"Error reading {ini_file}: {e}", exc_info=True)

    return sanitize_name(fld_dir.stem)


def build_path_map(root: Path):
    """Builds a mapping {folder_without_suffix: readable_name} for .ACT/.FLD."""
    name_map = {}
    for act in root.rglob("*.ACT"):
        name_map[act.with_suffix("")] = read_acct_name(act)
    for fld in root.rglob("*.FLD"):
        name_map[fld.with_suffix("")] = read_folder_name(fld)
    return name_map


def is_ole2_file(path: Path) -> bool:
    """Checks if the file has an OLE2 structure (a real Outlook .msg)."""
    try:
        with open(path, "rb") as f:
            header = f.read(8)
        # Signature for OLE2 Compound File Binary Format
        return header == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
    except Exception:
        return False


def convert_msg_to_eml(msg_path: Path, eml_path: Path):
    """Converts real Outlook MSG files or copies text-based emails."""
    try:
        # First, check if it's a genuine Outlook MSG file
        if is_ole2_file(msg_path):
            msg = extract_msg.Message(str(msg_path))
            eml_bytes = msg.as_bytes()
            with open(eml_path, "wb") as f:
                f.write(eml_bytes)
            #logging.info(f"CONVERTED (Outlook MSG): {msg_path}")
            return True

        # Otherwise, check if it’s a text-based EML file
        with open(msg_path, "rb") as f:
            data = f.read(4096)

        text = data.decode("utf-8", errors="ignore")
        if any(h in text for h in ("From:", "Subject:", "Content-Type:", "Return-Path:")):
            eml_path.write_text(text, encoding="utf-8", errors="ignore")
            #logging.info(f"COPIED (Text mail): {msg_path}")
            return True

        # Fallback: PMMail binary format – at least save raw content
        eml_path.write_bytes(data)
        logging.warning(f"UNCLEAR (Binary format, copied without conversion): {msg_path}")
        return True

    except Exception as e:
        logging.error(f"ERROR processing {msg_path}: {e}", exc_info=True)
        return False


# --------------------------------------------------------
# MAIN PROGRAM
# --------------------------------------------------------

def main():
    if not SOURCE_DIR.exists():
        print("Source directory does not exist.")
        return

    TARGET_DIR.mkdir(parents=True, exist_ok=True)
    logging.info("Starting conversion...")

    # Build name mapping for all .ACT and .FLD directories
    name_map = build_path_map(SOURCE_DIR)

    # Collect all .msg files recursively
    msg_files = [p for p in SOURCE_DIR.rglob("*") if p.suffix.lower() == ".msg"]
    if not msg_files:
        print("No .msg files found.")
        return

    ok_count, err_count = 0, 0

    for msg_file in tqdm(msg_files, desc="Converting .MSG files", unit="file"):
        rel_path = msg_file.relative_to(SOURCE_DIR)
        parts = list(rel_path.parts)

        for i in range(len(parts)):
            abs_path = SOURCE_DIR.joinpath(*rel_path.parts[:i + 1])
            if abs_path.with_suffix("") in name_map:
                parts[i] = name_map[abs_path.with_suffix("")]

        eml_rel = Path(*parts).with_suffix(".eml")
        eml_path = TARGET_DIR / eml_rel
        eml_path.parent.mkdir(parents=True, exist_ok=True)

        if convert_msg_to_eml(msg_file, eml_path):
            ok_count += 1
        else:
            err_count += 1

    print(f"\nDone. Successful: {ok_count}, Errors: {err_count}. Log: {LOG_FILE}")
    logging.info(f"Done. Successful: {ok_count}, Errors: {err_count}")

if __name__ == "__main__":
    main()
