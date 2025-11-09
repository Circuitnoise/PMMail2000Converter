# PMMail Archive Converter

This Python script converts legacy **PMMail 2000 email archives** into modern, portable **EML files**.  
It extracts folder and account names from `ACCT.INI` and `FOLDER.INI`, rebuilds the directory structure with human-readable names, and automatically converts all `.msg` files — whether they are real Outlook MSG files or text-based messages.

---

## 🧩 Features

- **Automatic detection and conversion** of Outlook `.msg` files to `.eml`
- **Parses PMMail structures** (`.ACT`, `.FLD`, `ACCT.INI`, `FOLDER.INI`)
- **Recursive directory processing** with a progress bar (`tqdm`)
- **File name sanitization** to remove invalid characters
- **Detailed logging** to `conversion_log.txt`
- **Robust fallback mechanisms** for incomplete or corrupted data

---

## 📦 Installation

### Prerequisites

- Python **3.9** or higher  
- macOS, Linux, or Windows  
- Write permissions for the target directory  

### 1️⃣ Clone the repository

```bash
git clone https://github.com/<USERNAME>/pmmail-converter.git
cd pmmail-converter
```

### 2️⃣ Create and activate a virtual environment (recommended)

```bash
python -m venv .venv
source .venv/bin/activate      # macOS/Linux
.venv\Scripts\activate         # Windows
```

### 3️⃣ Install dependencies

```bash
pip install -r requirements.txt
```

---

## ⚙️ Configuration

Edit the path variables at the top of the script to match your environment:

```python
SOURCE_DIR = Path("/path/to/PMMail-archive")
TARGET_DIR = Path("/path/to/output-folder")
```

- **SOURCE_DIR** – The root directory of your PMMail archive  
- **TARGET_DIR** – The folder where converted `.eml` files will be stored  

---

## ▶️ Usage

Run the script with Python:

```bash
python convert_pmmail_to_eml.py
```

You’ll see a progress bar during conversion:

```
Converting .MSG files: 100%|██████████| 243/243 [00:12<00:00, 20.0 file/s]
```

At the end, a summary appears:

```
Done. Successful: 238, Errors: 5. Log: /path/to/output/conversion_log.txt
```

---

## 🧠 How It Works

1. **Mapping PMMail metadata**  
   The script scans for `.ACT` and `.FLD` files and extracts human-readable account and folder names from `ACCT.INI` and `FOLDER.INI`.

2. **Recursive discovery**  
   It searches for all `.msg` files under the source directory.

3. **File conversion**  
   - Checks for **OLE2 signatures** (real Outlook MSG files)  
   - Uses `extract_msg` to convert them to EML  
   - Detects text-based emails and saves them directly  
   - Copies unknown binary formats as-is, logging a warning  

4. **Directory reconstruction**  
   The target folder structure mirrors the logical names from PMMail metadata.

5. **Logging**  
   All actions and errors are recorded in `conversion_log.txt`.

---

## 📄 Example Output

```
SOURCE_DIR: /Volumes/BackupC/Archiv/emailarchiv/blueprint Works/PMMail 2000
TARGET_DIR: /Volumes/BackupC/Archiv/emailarchiv/Convert

Starting conversion...
Converting .MSG files:  83%|████████▎ | 83/100 [00:06<00:01, 15.8 file/s]
Done. Successful: 98, Errors: 2. Log: /Volumes/BackupC/Archiv/emailarchiv/Convert/conversion_log.txt
```

---

## 🚀 Command-Line Usage

Once installed (via `pip install -e .`), the converter can be run from any terminal:

```bash
pmmail-convert
```

Optional arguments such as custom source or target paths can be added later if desired.

---

## 🧰 Troubleshooting

- **"Source directory does not exist"** → Check the `SOURCE_DIR` path  
- **Error reading ACCT.INI / FOLDER.INI** → The file might be corrupted; directory name is used as a fallback  
- **Unknown binary format** → File is copied as-is and noted in the log  

---

## 🪄 Developer Setup

To install in editable (dev) mode:

```bash
pip install -e .
```

Then run from anywhere:

```bash
pmmail-convert
```

To build a distributable package:

```bash
python -m build
```

---

## 🪪 License

This project is licensed under the **MIT License**.  
You are free to use, modify, and distribute it as long as the license notice remains included.
