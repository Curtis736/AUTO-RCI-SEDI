"""
Configuration globale pour pytest
Définit les fixtures partagées et la configuration des tests
"""

import sys
import os

# Ajouter le répertoire racine au chemin Python
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

# Configuration pytest
import pytest

@pytest.fixture(scope="session")
def project_root():
    """Retourne le chemin du répertoire racine du projet"""
    return root_dir

@pytest.fixture(scope="session")
def test_data_dir():
    """Retourne le chemin du dossier de données de test"""
    return os.path.join(root_dir, "tests", "fixtures", "test_data")

@pytest.fixture(scope="session")
def test_files_dir():
    """Retourne le chemin du dossier de fichiers de test"""
    return os.path.join(root_dir, "tests", "fixtures", "test_files")











