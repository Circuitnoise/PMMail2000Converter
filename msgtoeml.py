import os
import re
import shutil
import logging
from pathlib import Path
from tqdm import tqdm
import extract_msg
import struct

# --------------------------------------------------------
# KONFIGURATION
# Change these paths as needed
# --------------------------------------------------------

SOURCE_DIR = Path("/Volumes/blueprint Works/Blueprint Software Works/PMMail 2000")
TARGET_DIR = Path("/Volumes/Convert")
LOG_FILE = TARGET_DIR / "conversion_log.txt"

# --------------------------------------------------------
# LOGGING INITIALISIEREN
# --------------------------------------------------------
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# --------------------------------------------------------
# HILFSFUNKTIONEN
# --------------------------------------------------------

def sanitize_name(name: str) -> str:
    """Entfernt ungültige Dateisystemzeichen und trimmt Leerzeichen."""
    return re.sub(r'[\\/*?:"<>|]', "_", name.strip())


def read_acct_name(act_dir: Path) -> str:
    ini_file = act_dir / "ACCT.INI"
    if not ini_file.exists():
        return sanitize_name(act_dir.stem)

    try:
        raw = ini_file.read_bytes()
        # Alle Nullbytes durch Trenner ersetzen
        text = raw.replace(b"\x00", b"|").decode("latin1", errors="ignore")

        # Suche nach ACCTNAME und dem folgenden Block
        m = re.search(r"ACCTNAME\|+([^|]+)", text)
        if m:
            name = m.group(1).strip()
            if name:
                return sanitize_name(name)

        # Fallback: eventuell steht es auch ohne Trenner direkt dahinter
        m = re.search(r"ACCTNAME\s*([A-Za-z0-9@._\-\s]+)", text)
        if m:
            return sanitize_name(m.group(1).strip())

    except Exception as e:
        logging.error(f"Fehler beim Lesen von {ini_file}: {e}", exc_info=True)

    # Fallback auf Verzeichnisnamen
    return sanitize_name(act_dir.stem)




def read_folder_name(fld_dir: Path) -> str:
    """Extrahiert Ordnernamen aus FOLDER.INI – robust gegen Encoding-Fehler."""
    ini_file = fld_dir / "FOLDER.INI"
    if not ini_file.exists():
        logging.warning(f"FOLDER.INI fehlt in {fld_dir}")
        return sanitize_name(fld_dir.stem)

    try:
        raw = ini_file.read_bytes()
        text = raw.replace(b"\x00", b"").decode(errors="ignore").strip()

        # nicht druckbare Zeichen ersetzen
        cleaned = re.sub(r"[^\x20-\x7E]", "|", text)

        # 1️⃣ Klassisches Format ("Posteingangfi1")
        match = re.search(r"^([\wÄÖÜäöüß\s\-]+?)(?:fi|\ufb01)\d", cleaned)
        if match:
            return sanitize_name(match.group(1))

        # 2️⃣ Normalisiertes Ş-Format oder kaputte Encodings: alles bis Zahl/Semikolon
        match = re.search(r"^[!]?([A-Za-zÄÖÜäöüß\s\-]+?)(?=[0-9;|])", cleaned)
        if match:
            return sanitize_name(match.group(1))

        # 3️⃣ Fallback – erstes Segment vor Trenner
        match = re.search(r"^[!]?([^|;\r\n]+)", cleaned)
        if match:
            return sanitize_name(match.group(1))

        logging.warning(f"Kein gültiger FOLDERNAME gefunden in {ini_file}")

    except Exception as e:
        logging.error(f"Fehler beim Lesen von {ini_file}: {e}", exc_info=True)

    return sanitize_name(fld_dir.stem)


def build_path_map(root: Path):
    """Erzeugt Mapping {Ordner_ohne_Suffix: sprechender_Name} für .ACT/.FLD."""
    name_map = {}
    for act in root.rglob("*.ACT"):
        name_map[act.with_suffix("")] = read_acct_name(act)
    for fld in root.rglob("*.FLD"):
        name_map[fld.with_suffix("")] = read_folder_name(fld)
    return name_map

def is_ole2_file(path: Path) -> bool:
    """Prüft, ob Datei eine OLE2-Struktur hat (also echtes Outlook .msg)."""
    try:
        with open(path, "rb") as f:
            header = f.read(8)
        # Signatur für OLE2 Compound File Binary Format
        return header == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
    except Exception:
        return False

def convert_msg_to_eml(msg_path: Path, eml_path: Path):
    """Konvertiert echte Outlook-MSG-Dateien oder kopiert Textmails."""
    try:
        # Zuerst prüfen, ob echte Outlook-MSG-Datei
        if is_ole2_file(msg_path):
            msg = extract_msg.Message(str(msg_path))
            eml_bytes = msg.as_bytes()
            with open(eml_path, "wb") as f:
                f.write(eml_bytes)
            #logging.info(f"KONVERTIERT (Outlook MSG): {msg_path}")
            return True

        # Wenn keine echte MSG, prüfen ob textbasierte EML-Datei
        with open(msg_path, "rb") as f:
            data = f.read(4096)

        text = data.decode("utf-8", errors="ignore")
        if any(h in text for h in ("From:", "Subject:", "Content-Type:", "Return-Path:")):
            eml_path.write_text(text, encoding="utf-8", errors="ignore")
            #logging.info(f"KOPIERT (Textmail): {msg_path}")
            return True

        # Fallback: PMMail binärformat – wenigstens Rohinhalt sichern
        eml_path.write_bytes(data)
        logging.warning(f"UNKLAR (Binärformat, unkonvertiert kopiert): {msg_path}")
        return True

    except Exception as e:
        logging.error(f"FEHLER bei {msg_path}: {e}", exc_info=True)
        return False


# --------------------------------------------------------
# HAUPTPROGRAMM
# --------------------------------------------------------

def main():
    if not SOURCE_DIR.exists():
        print("Quellverzeichnis existiert nicht.")
        return

    TARGET_DIR.mkdir(parents=True, exist_ok=True)
    logging.info("Starte Konvertierung...")

    # Mapping für alle .ACT- und .FLD-Verzeichnisse
    name_map = build_path_map(SOURCE_DIR)

    # Alle .msg-Dateien rekursiv sammeln
    msg_files = [p for p in SOURCE_DIR.rglob("*") if p.suffix.lower() == ".msg"]
    if not msg_files:
        print("Keine .msg-Dateien gefunden.")
        return

    ok_count, err_count = 0, 0

    for msg_file in tqdm(msg_files, desc="Konvertiere .MSG-Dateien", unit="Datei"):
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

    print(f"\nFertig. Erfolgreich: {ok_count}, Fehler: {err_count}. Log: {LOG_FILE}")
    logging.info(f"Fertig. Erfolgreich: {ok_count}, Fehler: {err_count}")

if __name__ == "__main__":
    main()
