"""거래유형 그룹핑·월세 정렬축·중복 후보 경고에 대한 회귀 테스트."""

import csv
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parents[1]
TOOL = SKILL_DIR / "scripts" / "maemul_tool.py"
LISTING_TEMPLATE = SKILL_DIR / "assets" / "매물장-템플릿.csv"


def listing_headers() -> list[str]:
    with LISTING_TEMPLATE.open(encoding="utf-8-sig", newline="") as handle:
        return next(csv.reader(handle))


class RankingAndDedupTests(unittest.TestCase):
    def run_tool(self, *args: str, expect_ok: bool = True) -> subprocess.CompletedProcess[str]:
        result = subprocess.run(
            [sys.executable, str(TOOL), *args], text=True, capture_output=True, check=False
        )
        if expect_ok and result.returncode != 0:
            self.fail(f"tool failed: {result.stderr}\n{result.stdout}")
        return result

    def write_csv(self, directory: Path, rows: list[dict[str, str]]) -> Path:
        headers = listing_headers()
        path = directory / "매물.csv"
        with path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=headers)
            writer.writeheader()
            for row in rows:
                writer.writerow({header: row.get(header, "") for header in headers})
        return path

    def search(self, path: Path, criteria: dict[str, object]) -> dict[str, object]:
        result = self.run_tool("search", "--file", str(path), "--criteria-json", json.dumps(criteria))
        return json.loads(result.stdout)

    ALIASES = {"매매가": "매매가(만원)", "보증금": "보증금(만원)", "월세": "월세(만원)"}

    def base_row(self, **overrides: str) -> dict[str, str]:
        row = {
            "번호": "P001", "상태": "진행", "종류": "원룸", "거래": "월세",
            "지역": "서울 마포구", "동네": "신촌동", "전용(㎡)": "23.1",
            "반려": "?", "옵션": "?", "접수일": "2026-08-20",
        }
        row.update({self.ALIASES.get(key, key): value for key, value in overrides.items()})
        return row

    # --- 거래유형이 하드 조건이 아니면 유형끼리 묶는다 ---------------------

    def test_mixed_transactions_are_grouped_and_warned(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = self.write_csv(Path(raw), [
                self.base_row(번호="P003", 거래="월세", 보증금="3000", 월세="80"),
                self.base_row(번호="P002", 거래="전세", 보증금="28000"),
                self.base_row(번호="P001", 거래="매매", 매매가="35000"),
            ])
            payload = self.search(path, {"hard": [{"field": "지역", "op": "contains", "value": "서울"}], "soft": []})

        order = [item["row"]["거래"] for item in payload["matches"]]
        self.assertEqual(order, ["매매", "전세", "월세"], "거래유형끼리 묶여야 한다")
        self.assertTrue(payload["warnings"], "혼재 경고가 있어야 한다")
        self.assertEqual(payload["warnings"][0]["transactions"], ["매매", "전세", "월세"])

    def test_hard_transaction_condition_keeps_single_ranking(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = self.write_csv(Path(raw), [
                self.base_row(번호="P001", 거래="월세", 보증금="1000", 월세="90"),
                self.base_row(번호="P002", 거래="월세", 보증금="3000", 월세="50"),
            ])
            payload = self.search(path, {"hard": [{"field": "거래", "op": "eq", "value": "월세"}], "soft": []})

        self.assertEqual(payload["warnings"], [], "단일 거래유형이면 경고가 없어야 한다")

    def test_transaction_in_condition_still_groups_actual_mixed_results(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = self.write_csv(Path(raw), [
                self.base_row(번호="P002", 거래="월세", 보증금="3000", 월세="50"),
                self.base_row(번호="P001", 거래="전세", 보증금="28000"),
            ])
            payload = self.search(path, {
                "hard": [{"field": "거래", "op": "in", "value": ["전세", "월세"]}],
                "soft": [],
            })

        self.assertEqual([item["row"]["거래"] for item in payload["matches"]], ["전세", "월세"])
        self.assertEqual(payload["warnings"][0]["transactions"], ["전세", "월세"])

    # --- 월세는 보증금이 아니라 월세를 먼저 본다 ---------------------------

    def test_monthly_rent_sorts_by_rent_before_deposit(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            path = self.write_csv(Path(raw), [
                self.base_row(번호="P001", 보증금="1000", 월세="90"),
                self.base_row(번호="P002", 보증금="3000", 월세="50"),
                self.base_row(번호="P003", 보증금="500", 월세="70"),
            ])
            payload = self.search(path, {"hard": [{"field": "거래", "op": "eq", "value": "월세"}], "soft": []})

        self.assertEqual([item["row"]["번호"] for item in payload["matches"]], ["P002", "P003", "P001"])

    # --- 텍스트 필드에 크기 비교를 걸면 조용히 통과시키지 않는다 -----------

    def test_ordered_operator_on_text_field_is_rejected(self) -> None:
        sample = SKILL_DIR / "assets" / "매물장-샘플.csv"
        for operator, value in (("gte", "남"), ("lte", "남"), ("between", ["남", "서"])):
            with self.subTest(operator=operator):
                result = self.run_tool(
                    "search", "--file", str(sample),
                    "--criteria-json", json.dumps({"hard": [{"field": "향", "op": operator, "value": value}], "soft": []}, ensure_ascii=False),
                    expect_ok=False,
                )
                self.assertNotEqual(result.returncode, 0, "조용히 통과하면 안 된다")
                self.assertIn("크기 비교", json.loads(result.stdout)["error"]["message"])

    def test_ordered_operator_on_numeric_field_still_works(self) -> None:
        sample = SKILL_DIR / "assets" / "매물장-샘플.csv"
        payload = json.loads(self.run_tool(
            "search", "--file", str(sample),
            "--criteria-json", json.dumps({"hard": [{"field": "월세(만원)", "op": "lte", "value": 100}], "soft": []}, ensure_ascii=False),
        ).stdout)
        self.assertTrue(payload["ok"])

    # --- 같은 물건이 다시 들어오면 막지 않고 알린다 ------------------------

    def add(self, workbook: Path, record: dict[str, str]) -> dict[str, object]:
        sha = json.loads(self.run_tool("hash", "--file", str(workbook)).stdout)["sha256"]
        result = self.run_tool(
            "add", "--workbook", str(workbook), "--expected-sha", sha,
            "--record-json", json.dumps(record, ensure_ascii=False),
        )
        return json.loads(result.stdout)

    def test_duplicate_listing_is_stored_with_warning(self) -> None:
        record = {
            "종류": "아파트", "거래": "매매", "지역": "서울 구로구", "동네": "구로동",
            "단지명": "교육용A단지", "동호": "101-1001", "매매가(만원)": "35000", "전용(㎡)": "84.9",
        }
        with tempfile.TemporaryDirectory() as raw:
            workbook = Path(raw) / "매물장.xlsx"
            self.run_tool("init-workbook", "--workbook", str(workbook))
            self.add(workbook, record)
            payload = self.add(workbook, {**record, "매매가(만원)": "34500"})

        self.assertEqual(payload["changed"]["번호"], "P002", "중복이어도 저장은 된다")
        self.assertTrue(payload["warnings"], "중복 후보 경고가 있어야 한다")
        self.assertEqual(payload["warnings"][0]["candidates"][0]["item_id"], "P001")

    def test_distinct_listing_has_no_duplicate_warning(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            workbook = Path(raw) / "매물장.xlsx"
            self.run_tool("init-workbook", "--workbook", str(workbook))
            self.add(workbook, {
                "종류": "아파트", "거래": "매매", "지역": "서울 구로구", "동네": "구로동",
                "단지명": "교육용A단지", "동호": "101-1001", "매매가(만원)": "35000", "전용(㎡)": "84.9",
            })
            payload = self.add(workbook, {
                "종류": "아파트", "거래": "매매", "지역": "서울 구로구", "동네": "구로동",
                "단지명": "교육용A단지", "동호": "102-503", "매매가(만원)": "33000", "전용(㎡)": "59.8",
            })

        self.assertEqual(payload["warnings"], [])

    def test_duplicate_normalizes_numbers_and_whitespace_but_keeps_region(self) -> None:
        base = {
            "종류": "아파트", "거래": "매매", "지역": "서울 구로구", "동네": "구로동",
            "단지명": "교육 용 A단지", "동호": "101동 1001호", "매매가(만원)": "35000.0", "전용(㎡)": "84.0",
        }
        with tempfile.TemporaryDirectory() as raw:
            workbook = Path(raw) / "매물장.xlsx"
            self.run_tool("init-workbook", "--workbook", str(workbook))
            self.add(workbook, base)
            duplicate = self.add(workbook, {
                **base, "단지명": "교육용a단지", "동호": "101-1001", "매매가(만원)": "35000", "전용(㎡)": "84",
            })
            other_region = self.add(workbook, {
                **base, "지역": "경기 구로구", "단지명": "교육용a단지", "동호": "101-1001",
                "매매가(만원)": "35000", "전용(㎡)": "84",
            })

        self.assertTrue(duplicate["warnings"])
        self.assertEqual(other_region["warnings"], [])


if __name__ == "__main__":
    unittest.main()
