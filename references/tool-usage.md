# 로컬 도구 사용법

`<tool>`은 `<skill-dir>/scripts/maemul_tool.py`다. 명령은 JSON을 출력하고 실패하면 0이 아닌 종료코드를 반환한다.
Google 읽기·쓰기는 Google Drive 도구가 맡고 이 CLI는 프로필과 로컬 Excel을 관리한다.

## 프로필

```bash
python3 <tool> profile show
python3 <tool> profile list
python3 <tool> profile activate --name <name>
python3 <tool> profile set --name <name> --access google-sheet \
  --sheet-id <id-or-url> --connector <connector> --account <email> --activate
python3 <tool> profile set --name <name> --access local-xlsx \
  --workbook /절대경로/매물장.xlsx --activate
```

저장소 형식은 version 1이다. 지원되는 프로필의 필수 필드를 검사하며 Excel은 파일·두 탭·헤더까지 검증한다. 같은 이름은
`--replace` 없이는 거부한다. 격리 테스트는 최상위 `--profile-store /tmp/profiles.json`을 사용한다.

## Excel 생성

```bash
python3 <tool> create-excel --directory /절대경로/폴더
python3 <tool> init-workbook --workbook /절대경로/지정이름.xlsx
```

`create-excel`은 빈 파일 생성·검증·고유 프로필 저장·활성화를 한 번에 수행한다. 기본 이름 충돌은 `-2`, `-3`으로
피한다. `init-workbook`은 지정 경로에 표준 파일만 만들며 기존 파일을 덮어쓰지 않는다.

## 읽기

```bash
python3 <tool> validate --workbook /절대경로/매물장.xlsx --kind listing
python3 <tool> hash --file /절대경로/매물장.xlsx
python3 <tool> inspect --workbook /절대경로/매물장.xlsx --id P001
python3 <tool> search --workbook /절대경로/매물장.xlsx --criteria-json '<JSON>' --limit 10
python3 <tool> search --workbook /절대경로/매물장.xlsx --criteria-json '<JSON>' --include-hold
```

연산자는 `eq`, `ne`, `in`, `not-in`, `contains`, `lte`, `gte`, `between`이다. 텍스트 필드의
크기 비교와 숫자 필드의 `contains`는 거부한다. `matches`와 `needs_verification`은 별도 집계다. `--include-hold`를
사용하면 `match_status_counts`와 `verification_status_counts`로 진행·보류 건수를 따로 확인한다.

## 변경

```bash
python3 <tool> add --workbook /절대경로/매물장.xlsx --record-json '<객체-or-배열>' --expected-sha <sha>
python3 <tool> update --workbook /절대경로/매물장.xlsx --id P001 --changes-json '<객체>' --expected-sha <sha>
python3 <tool> complete --workbook /절대경로/매물장.xlsx --id P001 --expected-sha <sha>
python3 <tool> detail-upsert --workbook /절대경로/매물장.xlsx --id P001 \
  --changes-json '<객체>' --expected-sha <sha>
```

`add`는 최대 1,000건을 받는다. 불완전한 핵심값은 보류와 경고를 만들고 중복 지문은 저장을 막지 않고 후보만 알린다.
`update`도 핵심값이 미확인이 되면 보류로 강제한다. `detail-upsert`는 `매물` 탭에 없는 번호를 거부한다.
거래는 매매·전세·월세·`?`만 허용한다. 쓰기는 요청 셀만 바꾸고 백업 경로와 새 해시를 반환한다.

표준 28열 CSV는 `validate --file`과 `search --file`에서 읽기 전용 스냅샷으로만 지원한다.

Google Sheets 추가 전에는 최신 28열 행과 신규 자료를 결정적으로 준비할 수 있다. 이 명령은 시트에 쓰지 않는다.

```bash
python3 <tool> prepare-add --existing-json '<기존 매물 배열>' --record-json '<신규 매물 객체-or-배열>'
```

반환된 `prepared`의 번호·상태를 그대로 배치 쓰기에 사용하고, `warnings`의 보류·중복 후보를 사용자에게 전달한다.
