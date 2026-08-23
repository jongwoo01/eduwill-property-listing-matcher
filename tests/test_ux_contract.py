import unittest
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parents[1]


class UserExperienceContractTests(unittest.TestCase):
    def test_skill_loads_only_route_specific_references(self) -> None:
        skill = (SKILL_DIR / "SKILL.md").read_text(encoding="utf-8")
        self.assertIn("references/user-experience.md", skill)
        self.assertIn("한 요청에서 모든 참고 문서를 읽지 않는다", skill)
        self.assertIn("references/excel-workflow.md", skill)
        self.assertNotIn("첫 사용자 응답 전에", skill)

    def test_onboarding_offers_exactly_google_or_excel_and_accepts_chat(self) -> None:
        setup = (SKILL_DIR / "references" / "setup.md").read_text(encoding="utf-8")
        self.assertIn("매물장 준비 1/4", setup)
        self.assertIn("매물장 준비 2/4", setup)
        self.assertIn("매물장 준비 3/4", setup)
        self.assertIn("매물장 준비 완료", setup)
        self.assertIn("Google Sheets", setup)
        self.assertIn("이 컴퓨터의 Excel", setup)
        self.assertIn("채팅에 붙여넣은 매물", setup)
        self.assertIn("상태=보류", setup)
        self.assertIn("SQLite나 CSV를 선택지로 제시하지 않는다", setup)

    def test_user_copy_covers_setup_search_change_and_recovery(self) -> None:
        ux = (SKILL_DIR / "references" / "user-experience.md").read_text(encoding="utf-8")
        required_sections = [
            "## 매물장 준비 4단계",
            "## 검색 결과 안내",
            "## 추가·수정·완료 확인",
            "## 오류와 복구",
            "## 피해야 할 문구",
            "## 완료 전 UX 점검",
        ]
        for section in required_sections:
            self.assertIn(section, ux)
        self.assertIn("원본 파일: 변경하지 않음", ux)
        self.assertIn("[매물장 열기]", ux)
        self.assertIn("사용자가 할 다음 행동", (SKILL_DIR / "references" / "google-sheets-workflow.md").read_text(encoding="utf-8"))

    def test_every_request_starts_with_connection_check(self) -> None:
        skill = (SKILL_DIR / "SKILL.md").read_text(encoding="utf-8")
        google = (SKILL_DIR / "references" / "google-sheets-workflow.md").read_text(encoding="utf-8")
        self.assertIn("## 매번 먼저 할 일 — 연결 확인", skill)
        self.assertIn("get_spreadsheet_metadata", skill)
        self.assertIn("workbook_exists", skill)
        self.assertIn("get_spreadsheet_metadata", google)

    def test_add_keeps_duplicate_check_and_reread_even_without_confirmation(self) -> None:
        skill = (SKILL_DIR / "SKILL.md").read_text(encoding="utf-8")
        google = (SKILL_DIR / "references" / "google-sheets-workflow.md").read_text(encoding="utf-8")
        self.assertIn("번호 중복 검사 → 쓰기 → 재조회 검증", skill)
        self.assertIn("확인을 생략하라고 해도", google)

    def test_switching_ledgers_is_non_destructive_and_never_deletes_rows(self) -> None:
        skill = (SKILL_DIR / "SKILL.md").read_text(encoding="utf-8")
        setup = (SKILL_DIR / "references" / "setup.md").read_text(encoding="utf-8")
        ux = (SKILL_DIR / "references" / "user-experience.md").read_text(encoding="utf-8")
        tools = (SKILL_DIR / "references" / "tool-usage.md").read_text(encoding="utf-8")
        self.assertIn("profile activate", skill)
        self.assertIn("행 자체는 남긴다", skill)
        self.assertIn("## 새 매물장과 전환", setup)
        self.assertIn("기존 프로필과 원본은 그대로 남는다", setup)
        self.assertIn("이전 매물장으로 돌아가줘", ux)
        self.assertIn("행을 지우지 않고", ux)
        for removed in ("start-new", "cancel-new", "next-name", "local-csv", "init-ledger", "normalize-pair"):
            self.assertNotIn(removed, skill + setup + ux + tools)

    def test_excel_workflow_handles_large_files_without_database_branch(self) -> None:
        excel = (SKILL_DIR / "references" / "excel-workflow.md").read_text(encoding="utf-8")
        schema = (SKILL_DIR / "references" / "sheet-schema.md").read_text(encoding="utf-8")
        self.assertIn("15,000건 이상", excel)
        self.assertIn("읽기 전용 모드", excel)
        self.assertIn("SQLite는 대안으로 제시하지 않는다", excel)
        self.assertNotIn("SQLite로 옮기는", schema)

    def test_ui_metadata_is_human_facing_and_invocable(self) -> None:
        lines = (SKILL_DIR / "agents" / "openai.yaml").read_text(encoding="utf-8").splitlines()
        values = {}
        for line in lines:
            stripped = line.strip()
            if ": " in stripped:
                key, value = stripped.split(": ", 1)
                values[key] = value.strip('"')
        self.assertGreaterEqual(len(values["short_description"]), 25)
        self.assertLessEqual(len(values["short_description"]), 64)
        self.assertIn("$maemul-matching", values["default_prompt"])
        self.assertIn("매물장", values["short_description"])


if __name__ == "__main__":
    unittest.main()
