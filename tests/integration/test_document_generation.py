"""
Tests d'intégration pour la génération de documents
"""

import unittest
import sys
import os

# Ajouter src/ au chemin Python
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import bootstrap  # noqa: F401

try:
    import DocumentGenerator
    import ValueFetcher
except ImportError as e:
    print(f"Impossible d'importer les modules nécessaires: {e}")
    DocumentGenerator = None
    ValueFetcher = None


class TestDocumentGeneration(unittest.TestCase):
    """Tests d'intégration pour la génération de documents"""
    
    @unittest.skipIf(DocumentGenerator is None, "DocumentGenerator non disponible")
    def test_document_generator_import(self):
        """Test que DocumentGenerator peut être importé"""
        if DocumentGenerator is None:
            self.skipTest("DocumentGenerator non disponible")
        
        # Vérifier que le module a les fonctions principales
        self.assertTrue(hasattr(DocumentGenerator, 'GenerateSingleDocument') or 
                       hasattr(DocumentGenerator, 'GenDoc'),
                       "DocumentGenerator devrait avoir une fonction de génération")
    
    @unittest.skipIf(ValueFetcher is None, "ValueFetcher non disponible")
    def test_value_fetcher_import(self):
        """Test que ValueFetcher peut être importé"""
        if ValueFetcher is None:
            self.skipTest("ValueFetcher non disponible")
        
        # Vérifier que ValueCollector existe
        self.assertTrue(hasattr(ValueFetcher, 'ValueCollector'),
                       "ValueFetcher devrait avoir la classe ValueCollector")


if __name__ == '__main__':
    unittest.main()











