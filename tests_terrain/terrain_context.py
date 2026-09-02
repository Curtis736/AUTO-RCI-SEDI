"""
Contexte de test terrain : charge un vrai LT depuis test-unitaire/ et configure Settings.
"""

import json
import os
import re
import shutil

import Settings


PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROFILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lt_profile.json")


def _abs(path: str) -> str:
    if not path:
        return ""
    if os.path.isabs(path):
        return os.path.normpath(path)
    return os.path.normpath(os.path.join(PROJECT_ROOT, path))


def load_profile(profile_path: str = None) -> dict:
    profile_path = profile_path or PROFILE_PATH
    with open(profile_path, "r", encoding="utf-8") as f:
        profile = json.load(f)
    resolved = dict(profile)
    for key in (
        "root_work_dir",
        "measurement_file",
        "dichroic_root",
        "image_folder",
    ):
        if key in resolved and resolved[key]:
            resolved[key] = _abs(resolved[key])
    return resolved


def apply_profile_to_settings(profile: dict):
    Settings.LoadConfig("paths", reload=True)
    Settings.LoadConfig("fields", reload=True)

    Settings.SetConfigValue("fields", "LT", profile.get("LT", ""))
    Settings.SetConfigValue("fields", "num_plan", profile.get("num_plan"))
    Settings.SetConfigValue("fields", "form", profile.get("form", ""))
    Settings.SetConfigValue("fields", "out_file_category", profile.get("out_file_category", "RCI"))

    Settings.SetConfigValue("paths", "root_work_dir", profile.get("root_work_dir", ""))
    Settings.SetConfigValue("paths", "measurement_file", profile.get("measurement_file", ""))
    Settings.SetConfigValue("paths", "generator_out_path", profile.get("generator_out_path", "_TESTS_SORTIE"))
    if profile.get("dichroic_root"):
        Settings.SetConfigValue("paths", "dichroic_root", profile["dichroic_root"])
    if profile.get("dichroic_filename_hint"):
        Settings.SetConfigValue("paths", "dichroic_filename_hint", profile["dichroic_filename_hint"])


def backup_config() -> str:
    backup_dir = os.path.join(PROJECT_ROOT, "config", "_backup_avant_tests")
    os.makedirs(backup_dir, exist_ok=True)
    for name in ("paths.json", "fields.json"):
        src = os.path.join(PROJECT_ROOT, "config", name)
        if os.path.isfile(src):
            shutil.copy2(src, os.path.join(backup_dir, name))
    return backup_dir


def restore_config_from_backup(backup_dir: str):
    if not backup_dir or not os.path.isdir(backup_dir):
        return
    for name in ("paths.json", "fields.json"):
        src = os.path.join(backup_dir, name)
        dst = os.path.join(PROJECT_ROOT, "config", name)
        if os.path.isfile(src):
            shutil.copy2(src, dst)


def ensure_test_output_dir(profile: dict) -> str:
    out = os.path.join(profile["root_work_dir"], profile.get("generator_out_path", "_TESTS_SORTIE"))
    os.makedirs(out, exist_ok=True)
    return out


def scan_sn_in_folder(folder: str) -> list:
    if not folder or not os.path.isdir(folder):
        return []
    pattern = re.compile(r"SN\s*(\d+-\d+-\d+)", re.IGNORECASE)
    found = set()
    for root, dirs, files in os.walk(folder):
        for name in list(dirs) + list(files):
            for match in pattern.finditer(name):
                found.add(match.group(1))
    return sorted(found)
