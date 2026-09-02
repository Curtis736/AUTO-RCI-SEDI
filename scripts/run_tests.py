#!/usr/bin/env python3
"""
Script de lancement pour tous les tests unitaires AUTO_RCI
"""

import sys
import os
import subprocess

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TESTS_DIR = os.path.join(PROJECT_ROOT, "tests")
SRC_DIR = os.path.join(PROJECT_ROOT, "src")
sys.path.insert(0, os.path.join(TESTS_DIR))
sys.path.insert(0, SRC_DIR)
os.chdir(PROJECT_ROOT)

def run_test_file(test_file):
    """Exécute un fichier de test et retourne le résultat"""
    print(f"\n{'='*60}")
    print(f"EXÉCUTION DE {test_file}")
    print(f"{'='*60}")
    
    try:
        result = subprocess.run([sys.executable, test_file], 
                              capture_output=True, 
                              text=True, 
                              cwd=PROJECT_ROOT)
        
        print(result.stdout)
        if result.stderr:
            print("ERREURS:")
            print(result.stderr)
        
        return result.returncode == 0
    except Exception as e:
        print(f"Erreur lors de l'exécution de {test_file}: {e}")
        return False

def main():
    """Fonction principale"""
    print("=== LANCEUR DE TESTS UNITAIRES AUTO_RCI ===")
    print("Ce script exécute tous les tests unitaires disponibles")
    print("="*60)
    
    # Liste des fichiers de test dans le répertoire racine (anciens)
    root_test_files = [
        os.path.join(TESTS_DIR, "test_config.py"),
    ]
    
    tests_dir = TESTS_DIR
    test_files = []
    
    # Ajouter les tests du répertoire racine
    for test_file in root_test_files:
        if os.path.exists(test_file):
            test_files.append(test_file)
    
    # Ajouter les tests du dossier tests/unit/
    unit_tests_dir = os.path.join(tests_dir, "unit")
    if os.path.exists(unit_tests_dir):
        for file in os.listdir(unit_tests_dir):
            if file.startswith("test_") and file.endswith(".py"):
                test_files.append(os.path.join(unit_tests_dir, file))
    
    # Ajouter les tests du dossier tests/integration/
    integration_tests_dir = os.path.join(tests_dir, "integration")
    if os.path.exists(integration_tests_dir):
        for file in os.listdir(integration_tests_dir):
            if file.startswith("test_") and file.endswith(".py"):
                test_files.append(os.path.join(integration_tests_dir, file))
    
    # Vérifier que les fichiers existent
    existing_tests = []
    for test_file in test_files:
        if os.path.exists(test_file):
            existing_tests.append(test_file)
        else:
            print(f"ATTENTION: Le fichier {test_file} n'existe pas")
    
    if not existing_tests:
        print("Aucun fichier de test trouvé!")
        return False
    
    # Exécuter les tests
    success_count = 0
    total_tests = len(existing_tests)
    
    for test_file in existing_tests:
        if run_test_file(test_file):
            success_count += 1
            print(f"[OK] {test_file} - RÉUSSI")
        else:
            print(f"[FAIL] {test_file} - ÉCHOUÉ")
    
    # Résumé final
    print(f"\n{'='*60}")
    print(f"RÉSUMÉ FINAL")
    print(f"{'='*60}")
    print(f"Tests exécutés: {total_tests}")
    print(f"Tests réussis: {success_count}")
    print(f"Tests échoués: {total_tests - success_count}")
    
    if success_count == total_tests:
        print("[SUCCESS] TOUS LES TESTS SONT PASSÉS!")
        return True
    else:
        print("[WARNING] CERTAINS TESTS ONT ÉCHOUÉ")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
