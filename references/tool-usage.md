# 결정적 로컬 도구

`<tool>`은 `<skill-dir>/scripts/maemul_tool.py`다. 명령은 JSON을 출력하고 실패 시 0이 아닌 종료코드를 반환한다.
쓰기 명령은 로컬 `.xlsx`에만 쓴다. Google 시트의 읽기·쓰기는 연결된 Google Drive 도구가 맡고, 이 도구는 프로필만 관리한다.

## 프로필

```bash
python3 <tool> profile show
python3 <tool> profile list
python3 <tool> profile activate --name <name>
```

Google 프로필 — `--sheet-id`에는 ID 또는 시트 링크 전체를 넘겨도 된다. 링크면 ID를 추출하고, `--spreadsheet-url`과
ID가 다른 시트를 가리키면 저장을 거부한다:

```bash
python3 <tool> profile set --name <name> --access google-sheet \
  --sheet-id <id 또는 링크> --connector <connector> --account <account> \
  --listing-sheet 매물 --detail-sheet 매물상세 --activate
```

Excel 프로필 — 파일과 두 탭이 실제로 읽히는지 확인한 뒤 저장한다:

```bash
python3 <tool> profile set --name <name> --access local-xlsx \
  --workbook /절대경로/매물장.xlsx --listing-sheet 매물 --detail-sheet 매물상세 --activate
```

같은 이름이 이미 있으면 `--replace` 없이는 거부한다. 테스트에서는 최상위 옵션 `--profile-store /tmp/profiles.json`으로
실제 저장소를 피한다.

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
`between`이다. 출력의 `matches`는 확인된 일치, `needs_verification`은 하드 조건 값이 `?`인 후보,
`relaxations`는 하드 조건을 하나씩 뺐을 때의 실제 건수다.

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
`상태=보류`가 되고 `warnings`에 이유가 나온다. `--expected-sha`가 현재 파일과 다르면 쓰지 않는다.

## CSV 스냅샷(읽기 전용)

```bash
python3 <tool> validate --file /절대경로/매물.csv --kind listing
python3 <tool> search --file /절대경로/매물.csv --criteria-json '<JSON>' --limit 10
```

표준 28열 CSV는 읽기 전용으로만 쓴다. Google 시트 행이 많아 순위가 헷갈릴 때 임시 스냅샷으로 내려 순위를
계산하는 용도다. CSV에 쓰는 명령은 없다.
