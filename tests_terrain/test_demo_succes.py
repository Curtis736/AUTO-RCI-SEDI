"""
Tests de démonstration — doivent TOUS réussir.
Utilise le LT réel défini dans lt_profile.json.
"""

import os
import sys
import unittest

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from tests_terrain.terrain_context import apply_profile_to_settings, load_profile, scan_sn_in_folder

import DocumentGenerator
import Settings
import ValueFetcher


class TestDemoSucces(unittest.TestCase):
    """Cas nominaux sur données test-unitaire/."""

    profile = None

    @classmethod
    def setUpClass(cls):
        cls.profile = load_profile()
        apply_profile_to_settings(cls.profile)

    def test_01_profil_lt_valide(self):
        self.assertRegex(self.profile["LT"], r"^LT\d+$")

    def test_02_fichier_mesures_present(self):
        self.assertTrue(os.path.isfile(self.profile["measurement_file"]))

    def test_03_dossier_lt_present(self):
        self.assertTrue(os.path.isdir(self.profile["root_work_dir"]))

    def test_04_extraction_lt_depuis_chemin(self):
        from PathsInterface import extract_lt_from_path

        lt = extract_lt_from_path(self.profile["root_work_dir"])
        self.assertEqual(lt, self.profile["LT"])

    def test_05_image_sn_trouvee(self):
        folder = self.profile.get("image_folder")
        sn = self.profile.get("sample_sn")
        self.assertTrue(folder and sn)
        found, _all = DocumentGenerator.FindImagesMatchingSN(folder, sn)
        self.assertEqual(len(found), 1)
        self.assertTrue(found[0].upper().endswith(".JPG"))

    def test_06_normalisation_lt(self):
        self.assertTrue(ValueFetcher._lt_numbers_equal("LT2400182", "lt2400182"))
        self.assertTrue(ValueFetcher._lt_numbers_equal("LT2400182", "LT02400182"))

    def test_07_scan_sn_dossier_images(self):
        sns = scan_sn_in_folder(self.profile["image_folder"])
        self.assertIn(self.profile["sample_sn"], sns)

    def test_08_settings_appliques(self):
        self.assertEqual(
            os.path.normpath(Settings.GetConfigValueString("paths", "root_work_dir")),
            os.path.normpath(self.profile["root_work_dir"]),
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
