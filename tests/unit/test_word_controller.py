"""
Tests unitaires pour WordController
"""

import unittest
import sys
import os

# Ajouter le répertoire racine au chemin Python
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

try:
    import WordController
except ImportError as e:
    print(f"Impossible d'importer WordController: {e}")
    WordController = None


class TestWordController(unittest.TestCase):
    """Tests pour le module WordController"""
    
    def setUp(self):
        """Configuration avant chaque test"""
        if WordController:
            # Fermer Word s'il est déjà ouvert
            try:
                WordController.CloseWord()
            except:
                pass
    
    def tearDown(self):
        """Nettoyage après chaque test"""
        if WordController:
            # Fermer Word après chaque test
            try:
                WordController.CloseWord()
            except:
                pass
    
    @unittest.skipIf(WordController is None, "WordController non disponible")
    def test_open_word(self):
        """Test l'ouverture de Word"""
        if not WordController:
            self.skipTest("WordController non disponible")
        
        result = WordController.OpenWord()
        self.assertTrue(result, "L'ouverture de Word devrait réussir")
        
        # Vérifier que Word est bien ouvert en vérifiant directement la variable du module
        # __wordInstance dans un module Python n'utilise pas le name mangling (seulement pour les classes)
        # On peut y accéder via le dictionnaire du module
        word_instance = WordController.__dict__.get('__wordInstance')
        self.assertIsNotNone(word_instance, "L'instance Word devrait être créée après OpenWord()")
    
    @unittest.skipIf(WordController is None, "WordController non disponible")
    def test_close_word(self):
        """Test la fermeture de Word"""
        if not WordController:
            self.skipTest("WordController non disponible")
        
        # Ouvrir d'abord
        WordController.OpenWord()
        # Puis fermer
        result = WordController.CloseWord()
        self.assertTrue(result, "La fermeture de Word devrait réussir")
    
    @unittest.skipIf(WordController is None, "WordController non disponible")
    def test_clear_gencache(self):
        """Test la fonction de nettoyage du cache gencache"""
        if not WordController:
            self.skipTest("WordController non disponible")
        
        result = WordController._ClearWordGenCache()
        # La fonction devrait au moins s'exécuter sans erreur
        self.assertIsInstance(result, bool)


if __name__ == '__main__':
    unittest.main()

