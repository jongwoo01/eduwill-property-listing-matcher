# 결정적 로컬 도구

`<tool>`은 `<skill-dir>/scripts/maemul_tool.py`다. 명령은 JSON을 출력하고 실패 시 0이 아닌 종료코드를 반환한다.

## 프로필

```bash
python3 <tool> profile show
python3 <tool> profile list
python3 <tool> profile start-new
python3 <tool> profile cancel-new
python3 <tool> profile next-name --base 기본매물장
python3 <tool> profile activate --name <name>
```

Google 프로필:

```bash
python3 <tool> profile set --name <name> --access google-sheet \
  --sheet-id <id> --spreadsheet-url <url> --connector <connector> --account <account> \
  --listing-sheet 매물 --detail-sheet 매물상세 --activate
```

Excel 프로필:

```bash
python3 <tool> profile set --name <name> --access local-xlsx \
  --workbook /절대경로/매물장.xlsx --listing-sheet 매물 --detail-sheet 매물상세 --activate
```

`local-csv`는 기존 프로필 호환용으로만 유지한다. 신규 사용자에게 제시하지 않는다. 테스트에서는 실제 저장소 대신 최상위
옵션 `--profile-store /tmp/profiles.json`을 사용한다.

## Excel 생성·검증·검색

```bash
python3 <tool> init-workbook --workbook /절대경로/매물장.xlsx
python3 <tool> validate --workbook /절대경로/매물장.xlsx --kind listing
python3 <tool> validate --workbook /절대경로/매물장.xlsx --kind detail
python3 <tool> hash --file /절대경로/매물장.xlsx
python3 <tool> search --workbook /절대경로/매물장.xlsx \
  --criteria-json '<JSON>' --limit 10
```

다른 탭 이름이면 `--sheet <이름>`을 쓴다. 지원 연산자는 `eq`, `ne`, `in`, `not-in`, `contains`, `lte`, `gte`,
`between`이다. 출력의 `matches`는 확인된 일치, `needs_verification`은 하드 조건 값이 `?`인 후보다.

## Excel 변경과 채팅 일괄 추가

```bash
python3 <tool> inspect --workbook /절대경로/매물장.xlsx --id P001
python3 <tool> add --workbook /절대경로/매물장.xlsx \
  --record-json '{"종류":"아파트","거래":"전세"}' --expected-sha <sha>
python3 <tool> add --workbook /절대경로/매물장.xlsx \
  --record-json '[{"종류":"아파트"},{"거래":"전세","지역":"서울"}]' --expected-sha <sha>
python3 <tool> update --workbook /절대경로/매물장.xlsx --id P001 \
  --changes-json '{"보증금(만원)":"48000"}' --expected-sha <sha>
python3 <tool> complete --workbook /절대경로/매물장.xlsx --id P001 --expected-sha <sha>
python3 <tool> detail-upsert --workbook /절대경로/매물장.xlsx --id P001 \
  --changes-json '{"특약·메모":"확인 필요"}' --expected-sha <sha>
```

`add`는 객체 또는 최대 1,000개 객체 배열을 받는다. `null`은 `?`로 바뀐다. 핵심값이 불완전한 새 행은 자동으로
`상태=보류`가 되고 `warnings`에 이유가 나온다.

## 기존 CSV 이관·호환

```bash
python3 <tool> normalize-pair \
  --listing-input /원본/매물.csv --detail-input /원본/매물상세.csv \
  --output-dir /새로운/표준자료
python3 <tool> validate --file /절대경로/매물.csv --kind listing
python3 <tool> search --file /절대경로/매물.csv --criteria-json '<JSON>' --limit 10
```

`normalize-pair`는 원본과 기존 출력을 덮어쓰지 않는다. CSV 기반 `init-ledger`, `inspect`, `add`, `update`, `complete`,
`detail-upsert`도 이전 프로필 호환을 위해 유지된다.
