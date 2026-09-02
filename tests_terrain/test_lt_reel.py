"""
Tests terrain sur un vrai LT (données dans test-unitaire/).
"""

import os
import sys
import unittest

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from tests_terrain.terrain_context import (
    apply_profile_to_settings,
    load_profile,
    scan_sn_in_folder,
)

import DocumentGenerator
import Settings
import ValueFetcher


class TestLtReel(unittest.TestCase):
    profile = None

    @classmethod
    def setUpClass(cls):
        cls.profile = load_profile()
        apply_profile_to_settings(cls.profile)
        ValueFetcher.valueCollector.LT = cls.profile.get("LT", "")
        ValueFetcher.valueCollector.numPlan = cls.profile.get("num_plan", "")

    def test_profil_lt_charge(self):
        self.assertTrue(self.profile.get("LT", "").startswith("LT"))
        self.assertIn("test-unitaire", self.profile["root_work_dir"])

    def test_dossiers_lt_existants(self):
        self.assertTrue(
            os.path.isdir(self.profile["root_work_dir"]),
            f"Dossier LT introuvable : {self.profile['root_work_dir']}",
        )
        self.assertTrue(
            os.path.isfile(self.profile["measurement_file"]),
            f"Fichier mesures introuvable : {self.profile['measurement_file']}",
        )

    def test_images_dans_dossier_lt(self):
        image_folder = self.profile.get("image_folder")
        if not image_folder or not os.path.isdir(image_folder):
            self.skipTest("Dossier images non configuré dans lt_profile.json")

        sns = scan_sn_in_folder(image_folder)
        self.assertGreater(len(sns), 0, f"Aucun SN trouvé dans {image_folder}")

        sample_sn = self.profile.get("sample_sn", sns[0])
        found, _all = DocumentGenerator.FindImagesMatchingSN(image_folder, sample_sn)
        self.assertGreater(
            len(found),
            0,
            f"Aucune image pour SN{sample_sn} dans {image_folder}",
        )

    def test_find_sn_from_lt_excel(self):
        if not os.path.isfile(self.profile["measurement_file"]):
            self.skipTest("Fichier mesures absent")

        collector = ValueFetcher.ValueCollector()
        collector.LT = self.profile["LT"]
        collector.numPlan = self.profile.get("num_plan", "")

        try:
            if not collector.FetchValuesInExcel():
                self.skipTest("Excel COM indisponible ou fichier non lisible")
        except Exception as exc:
            self.skipTest(f"Excel non disponible sur ce poste : {exc}")

        lt = self.profile["LT"]
        sns = collector.FindSNFromLT(lt)
        if not sns:
            sns = collector._find_sn_from_work_dir(lt)

        self.assertGreater(
            len(sns),
            0,
            f"Aucun SN pour {lt} (Excel + repli dossier)",
        )

    def test_lt_synchronise_depuis_chemin(self):
        from PathsInterface import extract_lt_from_path

        lt = extract_lt_from_path(self.profile["root_work_dir"])
        self.assertEqual(lt, self.profile["LT"])

    def test_config_settings_coherente(self):
        apply_profile_to_settings(self.profile)
        self.assertEqual(
            Settings.GetConfigValueString("fields", "LT"),
            self.profile["LT"],
        )
        self.assertEqual(
            os.path.normpath(Settings.GetConfigValueString("paths", "root_work_dir")),
            os.path.normpath(self.profile["root_work_dir"]),
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
