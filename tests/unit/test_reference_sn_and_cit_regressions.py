"""
Tests de non-régression:
- collisions de SN entre références distinctes
- alias de tags CIT (FIN_CIT / fin_cit)
"""

import unittest
import sys
import os

# Ajouter le répertoire racine au chemin Python
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import DocumentGenerator
import ValueFetcher


class DummyContainer:
    def __init__(self, sn, ref):
        self.SN = sn
        self.tagAndValues = {"REF_SEDI": ref}


class TestReferenceSnAndCitRegressions(unittest.TestCase):
    def test_container_selection_prefers_reference_and_sn(self):
        c_157 = DummyContainer("144", "RCT-AGS23.157C")
        c_159 = DummyContainer("144", "RCT-AGS23.159C")

        strict_tokens = {
            DocumentGenerator._normalize_selection_token("RCT-AGS23.157C|144")
        }
        self.assertTrue(DocumentGenerator._container_matches_selection(c_157, strict_tokens))
        self.assertFalse(DocumentGenerator._container_matches_selection(c_159, strict_tokens))

    def test_container_selection_keeps_legacy_sn_only_behavior(self):
        c_157 = DummyContainer("0144", "RCT-AGS23.157C")
        c_159 = DummyContainer("144", "RCT-AGS23.159C")

        legacy_tokens = {
            DocumentGenerator._normalize_selection_token("144")
        }
        # En mode historique SN seul, les deux références restent valides.
        self.assertTrue(DocumentGenerator._container_matches_selection(c_157, legacy_tokens))
        self.assertTrue(DocumentGenerator._container_matches_selection(c_159, legacy_tokens))

    def test_cit_aliases_from_fin_cit(self):
        tag_map = {"FIN_CIT": "CIT-2024-001"}
        ValueFetcher.EnsureCitAliases(tag_map)
        self.assertEqual(tag_map.get("fin_cit"), "CIT-2024-001")

    def test_cit_aliases_from_lowercase_fin_cit(self):
        tag_map = {"fin_cit": "CIT-2024-002"}
        ValueFetcher.EnsureCitAliases(tag_map)
        self.assertEqual(tag_map.get("FIN_CIT"), "CIT-2024-002")


if __name__ == "__main__":
    unittest.main()

