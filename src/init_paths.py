"""
Initialise la racine du projet, le répertoire de travail et sys.path.
Doit être importé avant tout autre module applicatif.
"""
import os
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(PROJECT_ROOT, "src")

os.chdir(PROJECT_ROOT)
for path in (SRC_DIR, PROJECT_ROOT):
    if path not in sys.path:
        sys.path.insert(0, path)
