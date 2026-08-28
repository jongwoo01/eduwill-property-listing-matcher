import re
import unittest
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parents[1]


class UserExperienceContractTests(unittest.TestCase):
    def test_skill_reference_links_resolve_and_removed_ux_file_is_not_referenced(self) -> None:
        skill = (SKILL_DIR / "SKILL.md").read_text(encoding="utf-8")
        self.assertEqual("eduwill-property-listing-matcher", SKILL_DIR.name)
        self.assertIn("name: eduwill-property-listing-matcher", skill)
        references = set(re.findall(r"references/[a-z0-9-]+\.md", skill))
        self.assertGreaterEqual(len(references), 7)
        for reference in references:
            self.assertTrue((SKILL_DIR / reference).is_file(), reference)
        self.assertNotIn("references/user-experience.md", skill)
        self.assertFalse((SKILL_DIR / "references" / "user-experience.md").exists())

    def test_storage_docs_define_conditional_google_dependency_and_exact_tools(self) -> None:
        google = (SKILL_DIR / "references" / "google-sheets-workflow.md").read_text(encoding="utf-8")
        excel = (SKILL_DIR / "references" / "excel-workflow.md").read_text(encoding="utf-8")
        for tool in (
            "get_profile",
            "get_spreadsheet_metadata",
            "get_spreadsheet_range",
            "get_spreadsheet_cells",
            "import_spreadsheet",
            "batch_update_spreadsheet",
        ):
            self.assertIn(tool, google)
        self.assertIn("Google Drive 플러그인이나 로그인이 필요 없다", excel)
        self.assertIn("자동 전환하지 않는다", google)

    def test_docs_are_compact_and_have_no_duplicate_mermaid_flow(self) -> None:
        self.assertLessEqual(len((SKILL_DIR / "SKILL.md").read_text(encoding="utf-8").splitlines()), 90)
        self.assertLessEqual(len((SKILL_DIR / "README.md").read_text(encoding="utf-8").splitlines()), 80)
        markdown = "\n".join(path.read_text(encoding="utf-8") for path in SKILL_DIR.rglob("*.md"))
        self.assertNotIn("```mermaid", markdown)

    def test_ui_metadata_supports_both_storage_modes_without_unconditional_dependency(self) -> None:
        content = (SKILL_DIR / "agents" / "openai.yaml").read_text(encoding="utf-8")
        self.assertIn("Google", content)
        self.assertIn("Excel", content)
        self.assertIn('display_name: "손님 맞춤 매물 찾기"', content)
        self.assertIn("$eduwill-property-listing-matcher", content)
        self.assertNotIn("dependencies:", content)

    def test_readme_installs_the_matching_github_repository_and_skill_folder(self) -> None:
        readme = (SKILL_DIR / "README.md").read_text(encoding="utf-8")
        self.assertIn("jongwoo01/eduwill-property-listing-matcher.git", readme)
        self.assertIn("~/.codex/skills/eduwill-property-listing-matcher", readme)


if __name__ == "__main__":
    unittest.main()
