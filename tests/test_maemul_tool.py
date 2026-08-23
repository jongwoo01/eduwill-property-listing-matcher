import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
TOOL = SKILL_DIR / "scripts" / "maemul_tool.py"
SHEET_ID = "1rr3GyCsLuuQJuo9x9kKaM3GKqEI4fp_JIMZDLfscn7k"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit?gid=0#gid=0"


class MaemulToolRegressionTests(unittest.TestCase):
    def run_tool(self, *args: str, expect_ok: bool = True) -> subprocess.CompletedProcess[str]:
        result = subprocess.run(
            [sys.executable, str(TOOL), *args],
            text=True,
            capture_output=True,
            check=False,
        )
        if expect_ok and result.returncode != 0:
            self.fail(f"tool failed: {result.stderr}\n{result.stdout}")
        return result

    def set_google_profile(self, store: Path, name: str, sheet_id: str, **extra: str) -> dict[str, object]:
        args = [
            "--profile-store", str(store), "profile", "set", "--name", name,
            "--access", "google-sheet", "--sheet-id", sheet_id,
            "--connector", "google-drive", "--account", "user@example.com",
            "--listing-sheet", "매물", "--detail-sheet", "매물상세", "--activate",
        ]
        for key, value in extra.items():
            args += [f"--{key.replace('_', '-')}", value]
        return json.loads(self.run_tool(*args).stdout)

    # --- Google 프로필 ---------------------------------------------------

    def test_google_profile_persists_operational_sheet_fields(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            store = Path(temporary) / "profiles.json"
            profile = self.set_google_profile(store, "기본매물장", SHEET_ID)["profile"]
            self.assertEqual(SHEET_ID, profile["sheet_id"])
            self.assertEqual(f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit", profile["spreadsheet_url"])
            self.assertEqual("매물", profile["listing_sheet"])
            self.assertEqual("매물상세", profile["detail_sheet"])
            shown = json.loads(self.run_tool("--profile-store", str(store), "profile", "show").stdout)
            self.assertEqual(SHEET_ID, shown["profile"]["sheet_id"])

    def test_google_profile_accepts_pasted_link_as_sheet_id(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            store = Path(temporary) / "profiles.json"
            profile = self.set_google_profile(store, "링크매물장", SHEET_URL)["profile"]
            self.assertEqual(SHEET_ID, profile["sheet_id"])
            self.assertNotIn("/d/https://", profile["spreadsheet_url"])

    def test_google_profile_rejects_mismatched_id_and_link(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            store = Path(temporary) / "profiles.json"
            result = self.run_tool(
                "--profile-store", str(store), "profile", "set", "--name", "불일치",
                "--access", "google-sheet", "--sheet-id", "AAAA_bbbb-1234567890",
                "--spreadsheet-url", SHEET_URL,
                "--connector", "google-drive", "--account", "user@example.com",
                expect_ok=False,
            )
            self.assertNotEqual(0, result.returncode)
            self.assertIn("서로 다른 시트", result.stdout)
            self.assertFalse(store.exists())

    def test_google_profile_rejects_malformed_sheet_id(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            store = Path(temporary) / "profiles.json"
            result = self.run_tool(
                "--profile-store", str(store), "profile", "set", "--name", "깨진ID",
                "--access", "google-sheet", "--sheet-id", "이건 아이디가 아님!",
                "--connector", "google-drive", "--account", "user@example.com",
                expect_ok=False,
            )
            self.assertNotEqual(0, result.returncode)
            self.assertIn("형식이 올바르지 않습니다", result.stdout)

    # --- 프로필 전환·보존 ------------------------------------------------

    def test_switching_profiles_keeps_both_and_refuses_silent_overwrite(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            store = Path(temporary) / "profiles.json"
            self.set_google_profile(store, "기본매물장", SHEET_ID)
            second = self.set_google_profile(store, "두번째", "1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789abcdef")
            self.assertEqual("두번째", second["active_profile"])

            overwritten = self.run_tool(
                "--profile-store", str(store), "profile", "set", "--name", "기본매물장",
                "--access", "google-sheet", "--sheet-id", "1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789abcdef",
                "--connector", "google-drive", "--account", "user@example.com",
                expect_ok=False,
            )
            self.assertIn("이미 있습니다", overwritten.stdout)

            returned = json.loads(self.run_tool(
                "--profile-store", str(store), "profile", "activate", "--name", "기본매물장",
            ).stdout)
            self.assertEqual(SHEET_ID, returned["profile"]["sheet_id"])
            listed = json.loads(self.run_tool("--profile-store", str(store), "profile", "list").stdout)
            self.assertEqual(2, len(listed["profiles"]))

    def test_legacy_interrupted_onboarding_restores_previous_profile(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            store = Path(temporary) / "profiles.json"
            self.set_google_profile(store, "기본매물장", SHEET_ID)
            data = json.loads(store.read_text(encoding="utf-8"))
            data["active_profile"] = None
            data["onboarding"] = {"mode": "new-ledger", "previous_active_profile": "기본매물장"}
            store.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

            shown = json.loads(self.run_tool("--profile-store", str(store), "profile", "show").stdout)
            self.assertEqual("기본매물장", shown["active_profile"])
            self.assertNotIn("onboarding", shown)

    def test_legacy_csv_profile_is_reported_as_unsupported(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            store = Path(temporary) / "profiles.json"
            store.write_text(json.dumps({
                "version": 1,
                "active_profile": "옛날",
                "profiles": {"옛날": {"name": "옛날", "access": "local-csv", "listing_path": "/x/매물.csv"}},
            }, ensure_ascii=False), encoding="utf-8")
            result = self.run_tool("--profile-store", str(store), "profile", "show", expect_ok=False)
            self.assertIn("지원하지 않는 저장 방식", result.stdout)

    # --- 로컬 Excel --------------------------------------------------------

    def test_local_excel_profile_search_and_incomplete_chat_batch(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            workbook = root / "고객매물장.xlsx"
            store = root / "profiles.json"
            initialized = json.loads(self.run_tool("init-workbook", "--workbook", str(workbook)).stdout)

            added = json.loads(self.run_tool(
                "add", "--workbook", str(workbook),
                "--record-json", json.dumps([
                    {
                        "종류": "아파트", "거래": "전세", "지역": "서울 구로구",
                        "보증금(만원)": 50000, "전용(㎡)": 84,
                    },
                    {"종류": "빌라", "전용(㎡)": 72},
                ], ensure_ascii=False),
                "--expected-sha", initialized["sha256"],
            ).stdout)
            self.assertEqual(2, len(added["changed"]))
            self.assertEqual("진행", added["changed"][0]["상태"])
            self.assertEqual("보류", added["changed"][1]["상태"])
            self.assertEqual(1, len(added["warnings"]))
            self.assertTrue(Path(added["backup"]).is_file())

            criteria = {
                "hard": [
                    {"field": "거래", "op": "eq", "value": "전세"},
                    {"field": "보증금(만원)", "op": "lte", "value": 50000},
                    {"field": "전용(㎡)", "op": "gte", "value": 80},
                ],
                "soft": [],
            }
            searched = json.loads(self.run_tool(
                "search", "--workbook", str(workbook),
                "--criteria-json", json.dumps(criteria, ensure_ascii=False),
            ).stdout)
            self.assertEqual(1, searched["matches_count"])
            self.assertEqual("P001", searched["matches"][0]["row"]["번호"])
            self.assertEqual(3, len(searched["relaxations"]))

            detailed = json.loads(self.run_tool(
                "detail-upsert", "--workbook", str(workbook), "--id", "P001",
                "--changes-json", json.dumps({"특약·메모": "확인 필요"}, ensure_ascii=False),
                "--expected-sha", added["sha256"],
            ).stdout)
            updated = json.loads(self.run_tool(
                "update", "--workbook", str(workbook), "--id", "P001",
                "--changes-json", json.dumps({"보증금(만원)": 48000}, ensure_ascii=False),
                "--expected-sha", detailed["sha256"],
            ).stdout)
            inspected = json.loads(self.run_tool("inspect", "--workbook", str(workbook), "--id", "P001").stdout)
            self.assertEqual("48000", inspected["row"]["보증금(만원)"])
            self.assertEqual(updated["sha256"], inspected["sha256"])

            stale = self.run_tool(
                "complete", "--workbook", str(workbook), "--id", "P001",
                "--expected-sha", initialized["sha256"], expect_ok=False,
            )
            self.assertIn("확인 이후 변경", stale.stdout)

            completed = json.loads(self.run_tool(
                "complete", "--workbook", str(workbook), "--id", "P001",
                "--expected-sha", updated["sha256"],
            ).stdout)
            self.assertEqual("완료", completed["changed"]["after"]["상태"])
            after = json.loads(self.run_tool(
                "search", "--workbook", str(workbook),
                "--criteria-json", json.dumps(criteria, ensure_ascii=False),
            ).stdout)
            self.assertEqual(0, after["matches_count"])
            self.assertEqual(2, after["validated_rows"])

            profile = json.loads(self.run_tool(
                "--profile-store", str(store), "profile", "set",
                "--name", "로컬매물장", "--access", "local-xlsx",
                "--workbook", str(workbook), "--activate",
            ).stdout)["profile"]
            self.assertEqual("local-xlsx", profile["access"])
            self.assertEqual(str(workbook.resolve()), profile["workbook_path"])

            shown = json.loads(self.run_tool("--profile-store", str(store), "profile", "show").stdout)["profile"]
            self.assertTrue(shown["workbook_exists"])

    def test_csv_snapshot_is_read_only(self) -> None:
        sample = SKILL_DIR / "assets" / "매물장-샘플.csv"
        searched = json.loads(self.run_tool(
            "search", "--file", str(sample),
            "--criteria-json", json.dumps({"hard": [{"field": "거래", "op": "eq", "value": "매매"}], "soft": []}),
        ).stdout)
        self.assertEqual(1, searched["matches_count"])  # P001만; P004 보류, P005 완료 제외
        self.assertEqual(sample.read_bytes(), sample.read_bytes())
        mutated = self.run_tool("complete", "--file", str(sample), "--id", "P001", "--expected-sha", "x", expect_ok=False)
        self.assertNotEqual(0, mutated.returncode)

    def test_template_headers_match_tool_schema(self) -> None:
        import csv
        import importlib.util

        spec = importlib.util.spec_from_file_location("maemul_tool", TOOL)
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(module)

        def header(name: str) -> list[str]:
            with (SKILL_DIR / "assets" / name).open(encoding="utf-8-sig", newline="") as handle:
                return next(csv.reader(handle))

        self.assertEqual(list(module.LISTING_COLUMNS), header("매물장-템플릿.csv"))
        self.assertEqual(list(module.LISTING_COLUMNS), header("매물장-샘플.csv"))
        self.assertEqual(list(module.DETAIL_COLUMNS), header("매물상세-템플릿.csv"))


if __name__ == "__main__":
    unittest.main()
