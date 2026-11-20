import pathlib
import re
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]
VBA = ROOT / "src" / "vba"

class VbaPlacementBoundaryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.core = (VBA / "Mimaki_Template_Imposer_v2_4.bas").read_text(encoding="utf-8")
        cls.vector = (VBA / "modMimakiPlacementVector.bas").read_text(encoding="utf-8")
        cls.vector_export = (VBA / "modMimakiPlacementVector_CODE_ONLY.txt").read_text(encoding="utf-8")

    def test_tiff_route_stays_on_stable_processor(self):
        self.assertIn("PlaceSourceRangeIntoSlot(srcDoc, outDoc, printRange, outPageIndex, slot, inputOrientation, PRESERVE_SOURCE_PAGE_OFFSET, True)", self.core)
        self.assertIn("Set placedPrint = MimakiV24_PlaceVectorPrintRangeIntoSlot(", self.core)

    def test_vector_engine_isolated_from_core(self):
        self.assertNotRegex(self.core, r"(?im)^Public Function MimakiV24_PlaceVectorPrintRangeIntoSlot")
        self.assertRegex(self.vector, r"(?im)^Public Function MimakiV24_PlaceVectorPrintRangeIntoSlot")

    def test_cross_module_dependencies_are_public(self):
        names = [
            "ResolveEffectiveInputOrientation", "GetSourceCanvasSizeForSlot",
            "GetSourceCropCenter", "IsSourceCandidateShape",
            "CreateTemporarySourceCanvasCopy", "EnsureDocumentPages",
            "EnsureLayer", "IsSlotsLayerName", "MoveLayerAboveReference",
            "NormalizePastedObjectToSingleShape", "RotateShapeToSlotOrientation",
            "CenterShapeOnPoint", "ClearTemporaryRasterLayer",
        ]
        for name in names:
            self.assertRegex(self.core, rf"(?im)^Public (?:Sub|Function) {name}\b", name)

    def test_run_handler_declares_vector_option_under_option_explicit(self):
        run_handler = re.search(r"(?is)Private Sub cmdRun_Click\(\)(.*?)End Sub", (VBA / "frmMimakiImposerPanel.frm").read_text(encoding="utf-8"))
        self.assertIsNotNone(run_handler)
        self.assertRegex(run_handler.group(1), r"(?im)^\s*Dim preserveVector As Boolean\s*$")
        self.assertRegex(run_handler.group(1), r"(?im)^\s*preserveVector = chkPreserveVector.Value\s*$")

    def test_vector_function_returns_through_its_declared_name(self):
        self.assertIn("Set MimakiV24_PlaceVectorPrintRangeIntoSlot = pastedShape", self.vector)
        self.assertNotIn("Set PlaceVectorPrintRangeIntoSlot = pastedShape", self.vector)
        self.assertTrue(self.vector_export.replace("\\r\\n", "\\n").endswith(self.vector.replace("\\r\\n", "\\n")))

    def test_vector_procedures_assign_only_to_declared_names(self):
        pattern = re.compile(
            r"(?im)^(Public|Private) (Sub|Function) ([A-Za-z_]\\w*)\\((.*?)\\)(?: As [A-Za-z_]\\w+)?[ \\t]*\\r?\\n([\\s\\S]*?)^End \\2\\b"
        )
        for match in pattern.finditer(self.vector):
            proc_name, header, body = match.group(3), match.group(4), match.group(5)
            declared = {proc_name.lower()}
            declared.update(name.lower() for name in re.findall(r"\\b(?:ByVal|ByRef|Optional)\\s+([A-Za-z_]\\w+)", header, re.I))
            declared.update(name.lower() for name in re.findall(r"(?im)^\\s*Dim\\s+([A-Za-z_]\\w+)", body))
            for name in re.findall(r"(?im)^\\s*(?:Set\\s+)?([A-Za-z_]\\w+)\\s*=(?!=)", body):
                self.assertIn(name.lower(), declared, f"{proc_name} assigns to undeclared name {name}")

    def test_vector_layers_remain_separately_placed_and_cleaned(self):
        self.assertIn("BuildSourceLayerRange(srcDoc, srcRange, sourceLayerName)", self.vector)
        self.assertIn("outLayer.Paste", self.vector)
        self.assertIn("ClearTemporaryRasterLayer srcDoc.ActivePage, True", self.vector)
        self.assertIn("srcDoc.ClearSelection", self.vector)

    def test_v24_header_matches_panel_generation(self):
        self.assertIn("' Version: v2.4.0", self.core)
        self.assertIn("' Generated: 2026-10-02", self.core)
        self.assertIn('Caption         =   "Mimaki Imposer v2.4.0 / 2026-10-02"', (VBA / "frmMimakiImposerPanel.frm").read_text(encoding="utf-8"))

if __name__ == "__main__":
    unittest.main()
