import csv
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
TOOL = SKILL_DIR / "scripts" / "maemul_tool.py"

LEGACY_LISTING_COLUMNS = [
    "번호", "상태", "종류", "거래", "지역", "단지명", "동호", "매매가(만원)",
    "보증금(만원)", "월세(만원)", "관리비(만원)", "전용(㎡)", "방수", "욕실", "층", "총층",
    "향", "준공", "주차(대)", "반려", "옵션", "입주가능", "역도보(분)", "학군(분)",
]

LEGACY_DETAIL_COLUMNS = [
    "번호", "소유자", "연락처", "권리관계", "임대차현황", "용도지역", "건폐율·용적률",
    "시설상태", "비선호시설", "공시가격(만원)", "특약·메모", "접수일", "접수경로",
]


def write_csv(path: Path, columns: list[str], row: dict[str, str]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, lineterminator="\n")
        writer.writeheader()
        writer.writerow(row)


def read_rows(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open(encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        return reader.fieldnames or [], list(reader)


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

    def set_google_profile(self, store: Path, name: str, sheet_id: str) -> dict[str, object]:
        result = self.run_tool(
            "--profile-store", str(store), "profile", "set", "--name", name,
            "--access", "google-sheet", "--sheet-id", sheet_id,
            "--spreadsheet-url", f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit",
            "--connector", "google-drive", "--account", "user@example.com",
            "--listing-sheet", "매물", "--detail-sheet", "매물상세", "--activate",
        )
        return json.loads(result.stdout)

    def test_legacy_pair_normalizes_without_overwriting_source(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "source"
            output = root / "output"
            source.mkdir()
            listing = source / "매물.csv"
            detail = source / "매물상세.csv"
            write_csv(listing, LEGACY_LISTING_COLUMNS, {
                "번호": "P001", "상태": "진행", "종류": "아파트", "거래": "매매",
                "지역": "서울 구로구 항동", "단지명": "테스트단지", "동호": "101동 101호",
                "매매가(만원)": "50000", "보증금(만원)": "", "월세(만원)": "",
                "관리비(만원)": "20", "전용(㎡)": "84", "방수": "3", "욕실": "2",
                "층": "10", "총층": "20", "향": "남", "준공": "2020", "주차(대)": "1",
                "반려": "가능", "옵션": "없음", "입주가능": "협의", "역도보(분)": "8",
                "학군(분)": "12",
            })
            write_csv(detail, LEGACY_DETAIL_COLUMNS, {
                "번호": "P001", "소유자": "홍길동", "연락처": "010-0000-0000",
                "권리관계": "?", "임대차현황": "?", "용도지역": "?", "건폐율·용적률": "?",
                "시설상태": "양호", "비선호시설": "없음", "공시가격(만원)": "30000",
                "특약·메모": "", "접수일": "2026-08-17", "접수경로": "전화",
            })
            listing_before = listing.read_bytes()
            detail_before = detail.read_bytes()

            self.run_tool(
                "normalize-pair", "--listing-input", str(listing), "--detail-input", str(detail),
                "--output-dir", str(output),
            )

            listing_headers, listing_rows = read_rows(output / "매물.csv")
            detail_headers, detail_rows = read_rows(output / "매물상세.csv")
            self.assertEqual(28, len(listing_headers))
            self.assertEqual(20, len(detail_headers))
            self.assertEqual("서울 구로구", listing_rows[0]["지역"])
            self.assertEqual("항동", listing_rows[0]["동네"])
            self.assertEqual("2026-08-17", listing_rows[0]["접수일"])
            self.assertEqual("?", listing_rows[0]["초등도보(분)"])
            self.assertEqual("기존 학군(분): 12", detail_rows[0]["학군상세"])
            self.assertEqual("30000", detail_rows[0]["공시가격"])
            self.assertEqual("Y", listing_rows[0]["반려"])
            self.assertEqual("N", listing_rows[0]["옵션"])
            self.assertEqual(listing_before, listing.read_bytes())
            self.assertEqual(detail_before, detail.read_bytes())

            second = self.run_tool(
                "normalize-pair", "--listing-input", str(listing), "--detail-input", str(detail),
                "--output-dir", str(output), expect_ok=False,
            )
            self.assertNotEqual(0, second.returncode)
            self.assertIn("덮어쓰지 않습니다", second.stdout)

    def test_google_profile_persists_operational_sheet_fields(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            store = Path(temporary) / "profiles.json"
            profile = self.set_google_profile(store, "기본매물장", "sheet123")["profile"]
            self.assertEqual("sheet123", profile["sheet_id"])
            self.assertEqual("매물", profile["listing_sheet"])
            self.assertEqual("매물상세", profile["detail_sheet"])
            shown = self.run_tool("--profile-store", str(store), "profile", "show")
            self.assertEqual("sheet123", json.loads(shown.stdout)["profile"]["sheet_id"])

    def test_start_new_preserves_old_profile_and_cancel_restores_it(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            store = Path(temporary) / "profiles.json"
            self.set_google_profile(store, "기본매물장", "sheet-old")

            started = json.loads(self.run_tool(
                "--profile-store", str(store), "profile", "start-new",
            ).stdout)
            self.assertTrue(started["started"])
            self.assertIsNone(started["active_profile"])
            self.assertEqual("기본매물장", started["onboarding"]["previous_active_profile"])
            self.assertEqual("sheet-old", started["previous_profile"]["sheet_id"])

            resumed = json.loads(self.run_tool(
                "--profile-store", str(store), "profile", "show",
            ).stdout)
            self.assertEqual("new-ledger", resumed["onboarding"]["mode"])
            self.assertIsNone(resumed["profile"])

            cancelled = json.loads(self.run_tool(
                "--profile-store", str(store), "profile", "cancel-new",
            ).stdout)
            self.assertTrue(cancelled["cancelled"])
            self.assertEqual("기본매물장", cancelled["active_profile"])
            self.assertEqual("sheet-old", cancelled["profile"]["sheet_id"])

            listed = json.loads(self.run_tool(
                "--profile-store", str(store), "profile", "list",
            ).stdout)
            self.assertEqual(1, len(listed["profiles"]))
            self.assertIsNone(listed["onboarding"])

    def test_new_sheet_completion_keeps_old_profile_and_uses_unique_name(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            store = Path(temporary) / "profiles.json"
            self.set_google_profile(store, "기본매물장", "sheet-old")
            self.run_tool("--profile-store", str(store), "profile", "start-new")

            suggested = json.loads(self.run_tool(
                "--profile-store", str(store), "profile", "next-name", "--base", "기본매물장",
            ).stdout)
            self.assertEqual("기본매물장 2", suggested["name"])

            created = self.set_google_profile(store, suggested["name"], "sheet-new")
            self.assertEqual("기본매물장 2", created["active_profile"])
            self.assertEqual("기본매물장", created["profile"]["previous_profile"])
            self.assertEqual("sheet-old", created["previous_profile"]["sheet_id"])

            listed = json.loads(self.run_tool(
                "--profile-store", str(store), "profile", "list",
            ).stdout)
            self.assertEqual(2, len(listed["profiles"]))
            self.assertIsNone(listed["onboarding"])

            returned = json.loads(self.run_tool(
                "--profile-store", str(store), "profile", "activate", "--name", "기본매물장",
            ).stdout)
            self.assertEqual("sheet-old", returned["profile"]["sheet_id"])

    def test_start_new_rejects_overwriting_previous_profile_even_with_replace(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            store = Path(temporary) / "profiles.json"
            self.set_google_profile(store, "기본매물장", "sheet-old")
            self.run_tool("--profile-store", str(store), "profile", "start-new")

            overwritten = self.run_tool(
                "--profile-store", str(store), "profile", "set", "--name", "기본매물장",
                "--access", "google-sheet", "--sheet-id", "sheet-new",
                "--connector", "google-drive", "--account", "user@example.com",
                "--replace", "--activate", expect_ok=False,
            )
            self.assertNotEqual(0, overwritten.returncode)
            self.assertIn("기존 프로필을 덮어쓰지 않습니다", overwritten.stdout)

            restored = json.loads(self.run_tool(
                "--profile-store", str(store), "profile", "cancel-new",
            ).stdout)
            self.assertEqual("sheet-old", restored["profile"]["sheet_id"])

    def test_local_excel_profile_search_and_incomplete_chat_batch(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            workbook = root / "고객매물장.xlsx"
            store = root / "profiles.json"
            initialized = json.loads(self.run_tool(
                "init-workbook", "--workbook", str(workbook),
            ).stdout)

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
            inspected = json.loads(self.run_tool(
                "inspect", "--workbook", str(workbook), "--id", "P001",
            ).stdout)
            self.assertEqual("48000", inspected["row"]["보증금(만원)"])
            self.assertEqual(updated["sha256"], inspected["sha256"])

            profile = json.loads(self.run_tool(
                "--profile-store", str(store), "profile", "set",
                "--name", "로컬매물장", "--access", "local-xlsx",
                "--workbook", str(workbook), "--activate",
            ).stdout)["profile"]
            self.assertEqual("local-xlsx", profile["access"])
            self.assertEqual(str(workbook.resolve()), profile["workbook_path"])

            shown = json.loads(self.run_tool(
                "--profile-store", str(store), "profile", "show",
            ).stdout)["profile"]
            self.assertTrue(shown["workbook_exists"])


if __name__ == "__main__":
    unittest.main()
