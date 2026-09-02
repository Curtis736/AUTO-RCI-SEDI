"""
Tests de non-régression pour l'auto-remplissage de la mesure de fin CIT :
- dernière valeur + abs + arrondi 3 décimales
- match REF + SN (pas de collision)
- ne pas écraser une valeur déjà présente
"""

import os
import sys
import tempfile
import shutil
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import CitMeasurementFetcher as cit
import ValueFetcher


class DummyContainer:
    def __init__(self, sn, ref, cit_value=None):
        self.SN = sn
        self.undefinedValues = []
        self.tagAndValues = {"REF_SEDI": ref}
        if cit_value is not None:
            self.tagAndValues["FIN_CIT"] = cit_value
            self.tagAndValues["fin_cit"] = cit_value
            self.tagAndValues["CIT"] = cit_value
        else:
            self.undefinedValues.extend(["CIT", "FIN_CIT", "fin_cit"])


class TestCitMeasurementFetcher(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="cit_test_")
        self.thermal = os.path.join(self.tmp, "THERMAL_CYCLING")
        os.makedirs(self.thermal)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _write_data_traitee_xlsm(self, filename, headers, last_values, extra_rows=None):
        """
        Crée un .xlsx (openpyxl sans vba) avec feuille data_traitee.
        Structure alignée sur les fichiers terrain :
        - ligne 3 = en-têtes REF SN...
        - lignes suivantes = données ; dernière ligne = mesure de fin
        """
        try:
            import openpyxl
        except ImportError:
            self.skipTest("openpyxl non disponible")

        path = os.path.join(self.thermal, filename)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "data_traitee"
        # Lignes 1-2 factices
        ws.append([None] * (4 + len(headers)))
        ws.append([None] * (4 + len(headers)))
        # Ligne 3 : headers décalés de 5 colonnes (comme terrain)
        header_row = [None, None, None, None, None] + list(headers)
        ws.append(header_row)
        # Ligne 4 : titres canaux
        ws.append(["date", "heure", "temps", "temp", "wl"] + [f"CH{i}" for i in range(1, len(headers) + 1)])
        # Quelques lignes de données
        if extra_rows:
            for row_vals in extra_rows:
                ws.append([None, None, None, None, None] + list(row_vals))
        ws.append([None, None, None, None, None] + list(last_values))
        wb.save(path)
        wb.close()
        return path

    def test_format_abs_and_round(self):
        self.assertEqual(cit._format_cit_value(0.052432), "0.052")
        self.assertEqual(cit._format_cit_value(-0.003078), "0.003")
        self.assertEqual(cit._format_cit_value("0,0589"), "0.059")
        self.assertIsNone(cit._format_cit_value(None))
        self.assertIsNone(cit._format_cit_value(""))

    def test_header_matches_ref_and_sn(self):
        self.assertTrue(cit._header_matches("RCT°-AGS23.157 SN144", "144", "RCT-AGS23.157C"))
        self.assertTrue(cit._header_matches("RCT°-AGS23.157 SN144", "0144", "23.157"))
        self.assertFalse(cit._header_matches("RCT°-AGS23.159 SN144", "144", "RCT-AGS23.157C"))
        self.assertFalse(cit._header_matches("RCT°-AGS23.157 SN145", "144", "RCT-AGS23.157C"))

    def test_header_matches_dashed_sn(self):
        self.assertTrue(
            cit._header_matches("RETA-AGS24.134B SN25-20-01", "25-20-01", "RETA-AGS24.134B")
        )
        self.assertFalse(
            cit._header_matches("RETA-AGS24.134B SN25-20-02", "25-20-01", "RETA-AGS24.134B")
        )

    def test_build_cit_index_batch(self):
        work = os.path.join(
            os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
            "test-unitaire",
            "LT2500543",
        )
        if not os.path.isdir(work):
            self.skipTest(f"Données locales absentes: {work}")
        cit.clear_cit_index_cache()
        index = cit.build_cit_index(work)
        self.assertGreater(len(index), 0)
        v1 = cit._lookup_cit_in_index(index, "25-20-01", "RETA-AGS24.134B")
        self.assertIsNotNone(v1)
        # Lookup depuis l'index : pas de ré-ouverture fichier par SN
        v2 = cit.fetch_fin_cit_for_sn(work, "25-20-01", "RETA-AGS24.134B", index=index)
        self.assertEqual(v1, v2)

    def test_extract_last_value_from_xlsm(self):
        self._write_data_traitee_xlsm(
            "mesure_157.xlsm",
            headers=["RCT°-AGS23.157 SN144", "RCT°-AGS23.157 SN145", "RCT°-AGS23.159 SN144"],
            extra_rows=[[0.01, 0.02, 0.99], [0.02, 0.03, 0.98]],
            last_values=[0.052432, -0.0589, 0.0681],
        )
        # Renommer en .xlsx car openpyxl write sans macros ; le fetcher accepte .xlsx aussi
        src = os.path.join(self.thermal, "mesure_157.xlsm")
        dst = os.path.join(self.thermal, "mesure_157.xlsx")
        os.rename(src, dst)

        value = cit.fetch_fin_cit_for_sn(self.tmp, "144", "RCT°-AGS23.157C")
        self.assertEqual(value, "0.052")

        # Autre référence → autre canal
        value_159 = cit.fetch_fin_cit_for_sn(self.tmp, "144", "RCT°-AGS23.159C")
        self.assertEqual(value_159, "0.068")

        # abs sur valeur négative
        value_145 = cit.fetch_fin_cit_for_sn(self.tmp, "145", "RCT-AGS23.157")
        self.assertEqual(value_145, "0.059")

    def test_do_not_overwrite_existing_cit(self):
        self._write_data_traitee_xlsm(
            "mesure.xlsx",
            headers=["RCT-AGS23.157 SN144"],
            last_values=[0.052],
        )
        container = DummyContainer("144", "RCT-AGS23.157C", cit_value="0.999")
        filled = cit.apply_fin_cit_to_container(container, self.tmp)
        self.assertFalse(filled)
        self.assertEqual(container.tagAndValues["FIN_CIT"], "0.999")
        self.assertEqual(container.tagAndValues["CIT"], "0.999")

    def test_fill_empty_cit_aliases(self):
        self._write_data_traitee_xlsm(
            "mesure.xlsx",
            headers=["RCT-AGS23.157 SN144"],
            last_values=[0.052432],
        )
        container = DummyContainer("144", "RCT-AGS23.157C")
        filled = cit.apply_fin_cit_to_container(container, self.tmp)
        self.assertTrue(filled)
        self.assertEqual(container.tagAndValues["CIT"], "0.052")
        self.assertEqual(container.tagAndValues["FIN_CIT"], "0.052")
        self.assertEqual(container.tagAndValues["fin_cit"], "0.052")
        self.assertNotIn("CIT", container.undefinedValues)
        self.assertNotIn("FIN_CIT", container.undefinedValues)

    def test_parse_acceptance_filename_channels(self):
        channels = cit._parse_acceptance_filename_channels(
            "oha acceptance  23.157 142-143-144-145-vide-146.XLS"
        )
        self.assertEqual(channels[0]["sn"], "142")
        self.assertEqual(channels[0]["ref_hint"], "23.157")
        self.assertEqual(channels[2]["sn"], "144")
        self.assertIsNone(channels[4])  # vide
        self.assertEqual(channels[5]["sn"], "146")

    def test_ensure_cit_aliases_still_work(self):
        tag_map = {"FIN_CIT": "0.052"}
        ValueFetcher.EnsureCitAliases(tag_map)
        self.assertEqual(tag_map.get("fin_cit"), "0.052")


if __name__ == "__main__":
    unittest.main()
