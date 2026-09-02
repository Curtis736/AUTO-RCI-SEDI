"""
Tests unitaires pour FileFinder
"""

import unittest
import sys
import os

# Ajouter src/ au chemin Python
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import bootstrap  # noqa: F401

try:
    import FileFinder
except ImportError as e:
    print(f"Impossible d'importer FileFinder: {e}")
    FileFinder = None


class TestFileFinder(unittest.TestCase):
    """Tests pour le module FileFinder"""
    
    def test_get_path_length(self):
        """Test la fonction GetPathLength"""
        if not FileFinder:
            self.skipTest("FileFinder non disponible")
        
        # Test avec chemin relatif
        self.assertEqual(FileFinder.GetPathLength("a/b/c"), 3)
        self.assertEqual(FileFinder.GetPathLength("a"), 1)
        self.assertEqual(FileFinder.GetPathLength(""), 0)
    
    def test_find_in_folders_recursively(self):
        """Test la recherche récursive dans les dossiers"""
        if not FileFinder:
            self.skipTest("FileFinder non disponible")
        
        # Test avec un chemin valide (si possible)
        import tempfile
        with tempfile.TemporaryDirectory() as tmpdir:
            # Créer une structure de test
            os.makedirs(os.path.join(tmpdir, "test1", "test2"), exist_ok=True)
            
            import re
            pattern = re.compile("test1")
            results = FileFinder.FindInFoldersRecursively(tmpdir, pattern, maxDepth=2)
            
            self.assertIsNotNone(results, "Les résultats ne devraient pas être None")


if __name__ == '__main__':
    unittest.main()











