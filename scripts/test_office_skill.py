"""Host-side tests for summary delivery, safe HTML and read-only file handling.

The JS engine parity tests generate the JSON fixtures used here. The browser smoke
page separately executes the actual bridge in a Pyodide Web Worker.
"""
import asyncio
from copy import deepcopy
import hashlib
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
import sys

sys.dont_write_bytecode = True

ROOT = Path(__file__).resolve().parent.parent
SPEC = importlib.util.spec_from_file_location("office_skill", ROOT / "skills/office-diff/scripts/office_diff.py")
SKILL = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(SKILL)


class OfficeSkillTests(unittest.TestCase):
    def setUp(self):
        self.source = json.loads((ROOT / "demo/artifacts/skill-bundle/word-budget.json").read_text(encoding="utf-8"))

    def test_summary_preserves_coverage_and_paginates_without_missing_changes(self):
        ids, offset = [], 0
        while True:
            page = SKILL._build_summary(self.source, offset=offset, limit=1, max_chars=4000)
            self.assertLessEqual(len(SKILL._json(page)), 4000)
            self.assertEqual(page["coverage"], self.source["coverage"])
            self.assertEqual(page["documents"], self.source["documents"])
            ids.extend(change["id"] for change in page["changes"])
            if not page["hasMore"]:
                self.assertIsNone(page["nextOffset"])
                break
            self.assertGreater(page["nextOffset"], offset)
            offset = page["nextOffset"]
        self.assertEqual(ids, [change["id"] for change in self.source["changes"]])

    def test_large_values_are_explicitly_truncated_within_the_budget(self):
        source = deepcopy(self.source)
        source["changes"][0]["before"]["value"]["data"] = "中文😀" * 100000
        source["changes"][0]["after"]["value"]["data"] = "修改" * 100000
        page = SKILL._build_summary(source, limit=100, max_chars=4000)
        self.assertLessEqual(len(SKILL._json(page)), 4000)
        self.assertTrue(page["truncated"])
        value = page["changes"][0]["before"]["value"]["data"]
        self.assertTrue(value["truncated"])
        self.assertEqual(value["originalChars"], 300000)
        self.assertEqual(len(source["changes"][0]["before"]["value"]["data"]), 300000)

    def test_filters_have_independent_offsets_and_reject_invalid_paging(self):
        page = SKILL._build_summary(self.source, aspect="text", limit=1)
        self.assertTrue(all(change["aspect"] == "text" for change in page["changes"]))
        for kwargs in ({"offset": -1}, {"limit": 0}, {"max_chars": 10}, {"offset": 999}, {"change_ids": "change-001"}):
            with self.assertRaises(ValueError):
                SKILL._build_summary(self.source, **kwargs)

    def test_report_escapes_document_html_and_has_no_executable_script(self):
        source = deepcopy(self.source)
        source["changes"][0]["target"]["label"] = '<img src=x onerror="alert(1)">'
        source["changes"][0]["after"]["value"]["data"] = '</pre><script>fetch("https://bad.invalid")</script>'
        report = SKILL.render_report(source)
        self.assertNotIn("<script", report.lower())
        self.assertNotIn("<img", report.lower())
        self.assertIn("&lt;script&gt;", report)
        self.assertIn("default-src 'none'", report)

    def test_file_outputs_never_overwrite_inputs_or_existing_results(self):
        async def exercise():
            with tempfile.TemporaryDirectory() as folder:
                folder = Path(folder)
                before, after = folder / "old.docx", folder / "new.docx"
                before.write_bytes(b"old"); after.write_bytes(b"new")
                original = SKILL.compare_bytes
                async def fixture_engine(*_args):
                    return deepcopy(self.source)
                SKILL.compare_bytes = fixture_engine
                try:
                    result = await SKILL.compare_files(before, after, folder / "result")
                    self.assertEqual(before.read_bytes(), b"old")
                    self.assertEqual(after.read_bytes(), b"new")
                    full = json.loads(Path(result["artifacts"]["changes"]).read_text(encoding="utf-8"))
                    self.assertEqual(full, self.source)
                    self.assertEqual(SKILL.get_change(result["artifacts"]["changes"], "change-001"), self.source["changes"][0])
                    for destination in (before, after, folder / "result"):
                        with self.assertRaises(FileExistsError):
                            await SKILL.compare_files(before, after, destination)
                finally:
                    SKILL.compare_bytes = original
        asyncio.run(exercise())

    def test_bundle_digest_matches_manifest(self):
        assets = ROOT / "skills/office-diff/assets"
        manifest = json.loads((assets / "engine-manifest.json").read_text(encoding="utf-8"))
        self.assertEqual(hashlib.sha256((assets / "office-diff.bundle.js").read_bytes()).hexdigest(), manifest["bundleSha256"])


if __name__ == "__main__":
    unittest.main()
