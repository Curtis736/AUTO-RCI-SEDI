#!/usr/bin/env python3
"""
Configuration de test pour les tests unitaires AUTO_RCI
Ce module fournit des configurations de test par défaut
"""

import os
import tempfile
import json

class TestConfig:
    """Configuration de test pour éviter les dépendances externes"""
    
    @staticmethod
    def create_test_config_files():
        """Crée des fichiers de configuration temporaires pour les tests"""
        test_dir = tempfile.mkdtemp(prefix="auto_rci_test_")
        
        # Configuration des champs
        fields_config = {
            "LT": "LT2500706",
            "form": "F487 D",
            "out_file_category": "RCI",
            "generator_out_filename": "RETA-697-HOI-25.008_SN**SN**_RCI",
            "file_reaction_mode": "remplacer",
            "num_plan": "25.008",
            "TITREPLAN": "Plan de Test",
            "pdf_input_path": "/test/path",
            "pdf_necessary": None
        }
        
        # Configuration des chemins
        paths_config = {
            "root_work_dir": "/test/work",
            "measurement_file": "/test/measurements.xlsx",
            "template_file_path": "/test/template.docx",
            "generator_out_path": "/test/output"
        }
        
        # Configuration des images
        images_config = {
            "default_image_path": "/test/images",
            "image_formats": ["jpg", "png", "bmp"]
        }
        
        # Configuration du thème
        theme_config = {
            "default_font": "Arial",
            "default_size": 12,
            "colors": {
                "primary": "#000000",
                "secondary": "#666666"
            }
        }
        
        # Créer les fichiers
        config_files = {
            "fields.json": fields_config,
            "paths.json": paths_config,
            "images.json": images_config,
            "theme.json": theme_config
        }
        
        config_paths = {}
        for filename, config_data in config_files.items():
            filepath = os.path.join(test_dir, filename)
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=2, ensure_ascii=False)
            config_paths[filename] = filepath
        
        return test_dir, config_paths
    
    @staticmethod
    def cleanup_test_config(test_dir):
        """Nettoie les fichiers de configuration temporaires"""
        import shutil
        try:
            shutil.rmtree(test_dir)
        except Exception as e:
            print(f"Erreur lors du nettoyage: {e}")
    
    @staticmethod
    def get_test_data():
        """Retourne des données de test pour les CIT et autres mesures"""
        return {
            "containers": [
                {
                    "SN": "SN001",
                    "tagAndValues": {
                        "fin_cit": "CIT-2024-001",
                        "**SN**": "SN001",
                        "**LT**": "LT2500706",
                        "**RL1**": "-45.2",
                        "**IL_BEFORE**": "-0.3"
                    },
                    "tagAndSpecs": {
                        "**RL1**": {"condition": 1, "value": 45},  # >= 45
                        "**IL_BEFORE**": {"condition": 0, "value": 0.5}  # <= 0.5
                    }
                },
                {
                    "SN": "SN002", 
                    "tagAndValues": {
                        "fin_cit": "CIT-2024-002",
                        "**SN**": "SN002",
                        "**LT**": "LT2500706",
                        "**RL1**": "-50.1",
                        "**IL_BEFORE**": "-0.2"
                    },
                    "tagAndSpecs": {
                        "**RL1**": {"condition": 1, "value": 45},
                        "**IL_BEFORE**": {"condition": 0, "value": 0.5}
                    }
                }
            ],
            "excel_data": {
                "SN": ["SN001", "SN002", "SN003"],
                "fin_cit": ["CIT-2024-001", "CIT-2024-002", "CIT-2024-003"],
                "**LT**": ["LT2500706", "LT2500706", "LT2500706"],
                "**RL1**": ["-45.2", "-50.1", "-48.3"],
                "**IL_BEFORE**": ["-0.3", "-0.2", "-0.4"]
            }
        }
    
    @staticmethod
    def get_cit_test_cases():
        """Retourne des cas de test spécifiques pour les CIT"""
        return [
            {
                "name": "CIT normal",
                "input": "CIT-2024-001",
                "expected": "CIT-2024-001"
            },
            {
                "name": "CIT numérique",
                "input": "123456",
                "expected": "123456"
            },
            {
                "name": "CIT vide",
                "input": "",
                "expected": ""
            },
            {
                "name": "CIT avec espaces",
                "input": " CIT-2024-001 ",
                "expected": "CIT-2024-001"
            },
            {
                "name": "CIT None",
                "input": None,
                "expected": ""
            }
        ]


def setup_test_environment():
    """Configure l'environnement de test"""
    test_dir, config_paths = TestConfig.create_test_config_files()
    
    # Modifier les variables d'environnement pour pointer vers les fichiers de test
    os.environ['AUTO_RCI_TEST_CONFIG_DIR'] = test_dir
    
    return test_dir, config_paths


def cleanup_test_environment(test_dir):
    """Nettoie l'environnement de test"""
    TestConfig.cleanup_test_config(test_dir)
    if 'AUTO_RCI_TEST_CONFIG_DIR' in os.environ:
        del os.environ['AUTO_RCI_TEST_CONFIG_DIR']


if __name__ == "__main__":
    # Test de la configuration
    print("=== TEST DE LA CONFIGURATION ===")
    
    test_dir, config_paths = setup_test_environment()
    print(f"Répertoire de test créé: {test_dir}")
    print("Fichiers de configuration créés:")
    for name, path in config_paths.items():
        print(f"  - {name}: {path}")
    
    # Afficher les données de test
    test_data = TestConfig.get_test_data()
    print(f"\nDonnées de test disponibles:")
    print(f"  - {len(test_data['containers'])} containers")
    print(f"  - {len(test_data['excel_data'])} colonnes Excel")
    
    cit_cases = TestConfig.get_cit_test_cases()
    print(f"  - {len(cit_cases)} cas de test CIT")
    
    # Nettoyage
    cleanup_test_environment(test_dir)
    print(f"\nNettoyage terminé: {test_dir}")


















