#!/usr/bin/env python3
"""
Lance tous les tests unitaires + tests terrain sur un vrai LT (test-unitaire/).
"""

import os
import subprocess
import sys
import unittest

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TERRAIN_DIR = os.path.dirname(os.path.abspath(__file__))

if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

import Settings
from tests_terrain.terrain_context import (
    apply_profile_to_settings,
    backup_config,
    load_profile,
    restore_config_from_backup,
)


def _run_unittest_module(module_name: str) -> bool:
    loader = unittest.TestLoader()
    suite = loader.loadTestsFromName(module_name)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    return result.wasSuccessful()


def _run_unittest_package(start_dir: str, pattern: str = "test_*.py") -> bool:
    loader = unittest.TestLoader()
    suite = loader.discover(start_dir=start_dir, pattern=pattern, top_level_dir=PROJECT_ROOT)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    return result.wasSuccessful()


def _run_subprocess_test(path: str) -> bool:
    rel = os.path.relpath(path, PROJECT_ROOT)
    print(f"\n{'=' * 60}\nExécution : {rel}\n{'=' * 60}")
    proc = subprocess.run([sys.executable, path], cwd=PROJECT_ROOT)
    return proc.returncode == 0


def main() -> int:
    profile = load_profile()
    backup_dir = backup_config()

    print("=" * 60)
    print("TESTS TERRAIN AUTO_RCI — LT réel")
    print("=" * 60)
    print(f"Profil  : {profile.get('label', profile.get('LT'))}")
    print(f"LT      : {profile.get('LT')}")
    print(f"Dossier : {profile.get('root_work_dir')}")
    print(f"Mesures : {profile.get('measurement_file')}")
    print("=" * 60)

    apply_profile_to_settings(profile)

    results = []

    unit_dir = os.path.join(PROJECT_ROOT, "tests", "unit")
    if os.path.isdir(unit_dir):
        print("\n>>> Tests unitaires (tests/unit/)")
        results.append(("tests/unit", _run_unittest_package(unit_dir)))

    terrain_test = os.path.join(TERRAIN_DIR, "test_lt_reel.py")
    print("\n>>> Tests terrain LT réel (tests_terrain/)")
    results.append(("tests_terrain/test_lt_reel.py", _run_unittest_module("tests_terrain.test_lt_reel")))

    integration_test = os.path.join(PROJECT_ROOT, "tests", "integration", "test_document_generation.py")
    if os.path.isfile(integration_test):
        print("\n>>> Tests intégration")
        results.append(("tests/integration", _run_subprocess_test(integration_test)))

    restore_config_from_backup(backup_dir)
    Settings.LoadConfig("paths", reload=True)
    Settings.LoadConfig("fields", reload=True)

    print("\n" + "=" * 60)
    print("RÉSUMÉ")
    print("=" * 60)
    ok = 0
    for name, passed in results:
        status = "OK" if passed else "ÉCHEC"
        print(f"  [{status}] {name}")
        if passed:
            ok += 1

    print(f"\n{ok}/{len(results)} suites réussies")
    print("Config utilisateur restaurée depuis config/_backup_avant_tests/")
    return 0 if ok == len(results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
