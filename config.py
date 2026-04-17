"""
Configuration for Select by G Presentation Generator.
Paths are auto-detected for both Cowork sandbox and local environments.
"""
import os
from pathlib import Path

_this_file = Path(__file__).resolve()
try:
    COWORK_MNT = _this_file.parent.parent.parent
    if not COWORK_MNT.exists():
        COWORK_MNT = None
except Exception:
    COWORK_MNT = None

def find_gdrive_folder():
    """Find the mounted Google Drive folder."""
    if COWORK_MNT and COWORK_MNT.exists():
        for item in COWORK_MNT.iterdir():
            if item.is_dir() and 'Select' in item.name and 'G Group' in item.name:
                return item
    home = Path.home()
    for candidate in [
        Path('G:/My Drive/Select by G Group Pr\u00e9sentations'),
        Path('G:/My Drive') / 'Select by G Group Pr\u00e9sentations',
        home / 'Google Drive' / 'My Drive' / 'Select by G Group Pr\u00e9sentations',
        home / 'Google Drive' / 'Select by G Group Pr\u00e9sentations',
        home / 'Library' / 'CloudStorage' / 'GoogleDrive-ivanferrante1983@gmail.com' / 'My Drive' / 'Select by G Group Pr\u00e9sentations',
        home / 'Drive' / 'Select by G Group Pr\u00e9sentations',
    ]:
        if candidate.exists():
            return candidate
    fallback = Path(__file__).parent / 'output'
    fallback.mkdir(exist_ok=True)
    return fallback


def find_template():
    """Find the PPTX template file."""
    gdrive = find_gdrive_folder()
    if gdrive:
        for f in gdrive.iterdir():
            if 'Template' in f.name and f.suffix.lower() == '.pptx':
                return f
    if COWORK_MNT and COWORK_MNT.exists():
        for root, dirs, files in os.walk(str(COWORK_MNT)):
            for f in files:
                if 'Template' in f and f.endswith('.pptx') and 'Fournisseurs' in f:
                    return Path(root) / f
    raise FileNotFoundError(
        "Template PPTX not found. Please ensure 'Pr\u00e9sentation Fournisseurs Template EN.pptx' "
        "is in your Google Drive 'Select by G Group Pr\u00e9sentations' folder."
    )


GDRIVE_FOLDER = find_gdrive_folder()
TEMPLATE_PATH = None  # Lazy-loaded via find_template()

print(f"[Config] Google Drive folder: {GDRIVE_FOLDER}")
