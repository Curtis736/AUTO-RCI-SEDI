"""
Tests unitaires simples pour le module Log.
Le but est d'avoir un test rapide et lisible par tous.
"""

import unittest
import os
import sys

# Ajouter src/ au chemin Python
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import bootstrap  # noqa: F401

import Log


class TestLogModule(unittest.TestCase):
    """Tests faciles à lire pour vérifier la journalisation."""

    def setUp(self):
        # Sauvegarder l'état initial du module
        self._original_buffer = list(Log.logBuffer)
        self._original_callbacks = [list(callbacks) for callbacks in Log.logCallbacks]
        self._original_debug = Log.DEBUG_TO_PRINT

        # Préparer un environnement propre pour chaque test
        Log.logBuffer.clear()
        Log.logCallbacks = [[] for _ in range(len(Log.logCallbacks))]
        Log.DEBUG_TO_PRINT = False  # évite d'afficher les logs pendant les tests

    def tearDown(self):
        # Restaurer l'état initial du module
        Log.logBuffer[:] = self._original_buffer
        Log.logCallbacks = [list(callbacks) for callbacks in self._original_callbacks]
        Log.DEBUG_TO_PRINT = self._original_debug

    def test_message_appends_to_buffer(self):
        """Un appel à Log.Message() doit enregistrer le message dans le buffer."""
        Log.Message("Bonjour test")

        self.assertEqual(len(Log.logBuffer), 1)
        level, text = Log.logBuffer[0]
        self.assertEqual(level, Log.Lvl.MSG)
        self.assertIn("Bonjour test", text)

    def test_callback_receives_message(self):
        """Un callback enregistré doit recevoir le message envoyé."""
        received_messages = []

        def callback(message: str):
            received_messages.append(message.strip())

        Log.AddCallback(Log.Lvl.MSG, callback)
        Log.Message("Callback test")

        self.assertEqual(len(received_messages), 1)
        self.assertTrue(received_messages[0].endswith("Callback test"))


if __name__ == "__main__":
    unittest.main()








