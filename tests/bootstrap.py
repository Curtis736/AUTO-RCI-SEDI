#!/usr/bin/env python3
"""Configure sys.path et cwd pour les tests AUTO_RCI."""
import os
import sys

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(PROJECT_ROOT, "src")

os.chdir(PROJECT_ROOT)
for path in (SRC_DIR, PROJECT_ROOT):
    if path not in sys.path:
        sys.path.insert(0, path)
