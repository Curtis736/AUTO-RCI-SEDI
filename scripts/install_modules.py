#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script d'installation automatique de tous les modules nécessaires pour AUTO_RCI_COMPLET
"""

import subprocess
import sys
import os

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(PROJECT_ROOT, "src")
sys.path.insert(0, SRC_DIR)
os.chdir(PROJECT_ROOT)

def run_command(command, description):
    """Exécute une commande et affiche le résultat"""
    print(f"\n{'='*50}")
    print(f"{description}")
    print(f"{'='*50}")
    try:
        result = subprocess.run(
            command,
            shell=True,
            check=True,
            capture_output=True,
            text=True
        )
        if result.stdout:
            print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"ERREUR lors de {description}:")
        print(e.stderr)
        return False

def check_python_version():
    """Vérifie la version de Python"""
    version = sys.version_info
    print(f"Python {version.major}.{version.minor}.{version.micro} détecté")
    if version.major < 3 or (version.major == 3 and version.minor < 7):
        print("ATTENTION: Python 3.7 ou supérieur est recommandé")
    return True

def install_modules():
    """Installe tous les modules nécessaires"""
    print("\n" + "="*50)
    print("INSTALLATION DES MODULES POUR AUTO_RCI_COMPLET")
    print("="*50)
    
    # Vérifier la version de Python
    check_python_version()
    
    # Mettre à jour pip
    if not run_command(
        f"{sys.executable} -m pip install --upgrade pip",
        "Mise à jour de pip"
    ):
        print("ATTENTION: Échec de la mise à jour de pip, continuation...")
    
    # Vérifier si requirements.txt existe
    requirements_file = os.path.join(PROJECT_ROOT, "requirements.txt")
    if os.path.exists(requirements_file):
        print(f"\nInstallation depuis {requirements_file}...")
        if not run_command(
            f"{sys.executable} -m pip install -r {requirements_file}",
            f"Installation des modules depuis {requirements_file}"
        ):
            print("ERREUR: Échec de l'installation depuis requirements.txt")
            return False
    else:
        print(f"ATTENTION: {requirements_file} non trouvé, installation manuelle...")
        # Liste des modules à installer
        modules = [
            "pywin32",
            "Pillow",
            "numpy",
            "python-docx",
            "docx2pdf",
            "pdf2image",
            "pypdf",
            "openpyxl",
        ]
        
        for module in modules:
            if not run_command(
                f"{sys.executable} -m pip install {module}",
                f"Installation de {module}"
            ):
                print(f"ATTENTION: Échec de l'installation de {module}")
    
    # Installer les modules optionnels (non bloquants)
    print("\n" + "="*50)
    print("INSTALLATION DES MODULES OPTIONNELS (tests)")
    print("="*50)
    optional_modules = ["pytest"]
    for module in optional_modules:
        run_command(
            f"{sys.executable} -m pip install {module}",
            f"Installation optionnelle de {module}"
        )
    
    # Vérifier l'installation
    print("\n" + "="*50)
    print("VÉRIFICATION DE L'INSTALLATION")
    print("="*50)
    
    modules_to_check = [
        "pywin32",
        "Pillow",
        "numpy",
        "python-docx",
        "docx2pdf",
        "pdf2image",
        "pypdf",
        "openpyxl",
    ]
    
    installed_modules = []
    failed_modules = []
    
    for module in modules_to_check:
        try:
            result = subprocess.run(
                f"{sys.executable} -m pip show {module}",
                shell=True,
                capture_output=True,
                text=True
            )
            if result.returncode == 0:
                installed_modules.append(module)
                print(f"✓ {module} - Installé")
            else:
                failed_modules.append(module)
                print(f"✗ {module} - NON installé")
        except Exception as e:
            failed_modules.append(module)
            print(f"✗ {module} - Erreur de vérification: {e}")
    
    # Résumé
    print("\n" + "="*50)
    print("RÉSUMÉ")
    print("="*50)
    print(f"Modules installés avec succès: {len(installed_modules)}/{len(modules_to_check)}")
    
    if failed_modules:
        print(f"\nModules non installés: {', '.join(failed_modules)}")
        print("Vous pouvez les installer manuellement avec:")
        print(f"  {sys.executable} -m pip install {' '.join(failed_modules)}")
        return False
    else:
        print("\n✓ Tous les modules ont été installés avec succès!")

    print("\n" + "="*50)
    print("INSTALLATION POPPLER (PDF → IMAGE)")
    print("="*50)
    try:
        import poppler_setup
        poppler_bin = poppler_setup.ensure_poppler_installed()
        if poppler_bin:
            print(f"✓ Poppler prêt : {poppler_bin}")
        else:
            print("✗ Poppler non installé (connexion Internet requise au premier lancement)")
    except Exception as e:
        print(f"✗ Erreur installation Poppler : {e}")

    return True if not failed_modules else False

if __name__ == "__main__":
    try:
        success = install_modules()
        if success:
            print("\n" + "="*50)
            print("INSTALLATION TERMINÉE AVEC SUCCÈS!")
            print("="*50)
            sys.exit(0)
        else:
            print("\n" + "="*50)
            print("INSTALLATION TERMINÉE AVEC DES ERREURS")
            print("Veuillez vérifier les messages ci-dessus")
            print("="*50)
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n\nInstallation annulée par l'utilisateur")
        sys.exit(1)
    except Exception as e:
        print(f"\n\nERREUR INATTENDUE: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
