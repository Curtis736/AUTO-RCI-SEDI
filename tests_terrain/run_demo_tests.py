#!/usr/bin/env python3
"""
Démo tests : exécute d'abord les tests qui doivent RÉUSSIR,
puis les tests qui doivent ÉCHOUER (échecs attendus).
"""

import os
import sys
import unittest

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import Settings
from tests_terrain.terrain_context import (
    apply_profile_to_settings,
    backup_config,
    load_profile,
    restore_config_from_backup,
)


def _run_suite(module_name: str) -> tuple[bool, int, int, int]:
    loader = unittest.TestLoader()
    suite = loader.loadTestsFromName(module_name)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    total = result.testsRun
    failures = len(result.failures) + len(result.errors)
    passed = total - failures
    return result.wasSuccessful(), passed, failures, total


def main() -> int:
    profile = load_profile()
    backup_dir = backup_config()
    apply_profile_to_settings(profile)

    print("=" * 60)
    print("DÉMO TESTS AUTO_RCI")
    print("=" * 60)
    print(f"LT réel : {profile.get('LT')} — {profile.get('label', '')}")
    print("=" * 60)

    print("\n### ÉTAPE 1 / 2 — Tests qui DOIVENT réussir ###\n")
    ok_succes, pass_s, fail_s, total_s = _run_suite("tests_terrain.test_demo_succes")

    print("\n### ÉTAPE 2 / 2 — Tests qui DOIVENT échouer (démo) ###\n")
    ok_echec, pass_e, fail_e, total_e = _run_suite("tests_terrain.test_demo_echec")
    echecs_attendus = fail_e == total_e and total_e > 0 and pass_e == 0

    restore_config_from_backup(backup_dir)
    Settings.LoadConfig("paths", reload=True)
    Settings.LoadConfig("fields", reload=True)

    print("\n" + "=" * 60)
    print("RÉSUMÉ DÉMO")
    print("=" * 60)
    print(f"  Etape 1 (succes attendus) : {pass_s}/{total_s} OK  -> {'OK' if ok_succes else 'ECHEC'}")
    print(
        f"  Etape 2 (echecs attendus) : {fail_e}/{total_e} echoues  -> "
        f"{'OK (demo valide)' if echecs_attendus else 'PROBLEME (des tests auraient du echouer)'}"
    )
    if pass_e > 0:
        print(f"    Attention : {pass_e} test(s) ont réussi alors qu'un échec était attendu")
    print("=" * 60)

    demo_ok = ok_succes and echecs_attendus
    if demo_ok:
        print("\nDémo complète : la suite détecte correctement succès ET échecs.")
        return 0
    print("\nDémo incomplète : vérifiez les résultats ci-dessus.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
