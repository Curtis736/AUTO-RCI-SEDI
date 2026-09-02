"""
Tests de démonstration — doivent TOUS échouer volontairement.
Sert à vérifier que la suite de tests détecte bien les anomalies.
"""

import os
import sys
import unittest

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from tests_terrain.terrain_context import apply_profile_to_settings, load_profile

import DocumentGenerator
import ValueFetcher


class TestDemoEchec(unittest.TestCase):
    """Cas d'erreur simulés — échecs attendus."""

    profile = None

    @classmethod
    def setUpClass(cls):
        cls.profile = load_profile()
        apply_profile_to_settings(cls.profile)

    def test_01_chemin_mesures_inexistant(self):
        faux_chemin = os.path.join(PROJECT_ROOT, "test-unitaire", "FICHIER_QUI_N_EXISTE_PAS.xlsx")
        self.assertTrue(os.path.isfile(faux_chemin), "Ce test doit échouer : fichier inexistant")

    def test_02_lt_inexistant_dans_excel(self):
        collector = ValueFetcher.ValueCollector()
        collector.LT = self.profile["LT"]
        collector.FetchValuesInExcel()
        sns = collector.FindSNFromLT("LT9999999")
        self.assertGreater(len(sns), 0, "Ce test doit échouer : LT9999999 introuvable")

    def test_03_image_sn_inexistant(self):
        folder = self.profile.get("image_folder")
        self.assertIsNotNone(folder)
        found, _all = DocumentGenerator.FindImagesMatchingSN(folder, "99-99-99")
        self.assertGreater(len(found), 0, "Ce test doit échouer : SN99-99-99 absent")

    def test_04_dossier_lt_faux(self):
        faux_lt = os.path.join(PROJECT_ROOT, "test-unitaire", "LT9999999")
        self.assertTrue(os.path.isdir(faux_lt), "Ce test doit échouer : dossier LT fictif")

    def test_05_comparaison_lt_incorrecte(self):
        self.assertTrue(
            ValueFetcher._lt_numbers_equal("LT2400182", "LT2501492"),
            "Ce test doit échouer : LTs différents considérés égaux",
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
