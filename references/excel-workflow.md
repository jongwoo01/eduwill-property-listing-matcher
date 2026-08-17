# 로컬 Excel 흐름

사용자가 `local-xlsx`를 선택했을 때만 읽는다. 한 `.xlsx` 파일의 `매물`·`매물상세` 탭이 운영 원본이다.

## 연결·생성

1. 사용자가 지정한 폴더 또는 파일 경로를 절대 경로로 확인한다.
2. 새 파일이면 `init-workbook`으로 만들고, 같은 이름이 있으면 덮어쓰지 않는다.
3. 기존 파일이면 두 탭과 헤더를 `validate --workbook`으로 검증한다.
4. `profile set --access local-xlsx`로 파일 경로와 탭 이름을 저장한다.

폴더가 이동되거나 파일이 사라지면 새 파일을 자동 생성하지 말고 새 위치를 받는다. 동기화 폴더도 가능하지만 여러 사람이
동시에 편집한다면 Google Sheets를 안내한다.

## 검색

```bash
python3 <tool> search --workbook /절대경로/매물장.xlsx \
  --sheet 매물 --criteria-json '<JSON>' --limit 10
```

검색은 openpyxl의 읽기 전용 모드로 최신 파일을 한 번 순회한다. 15,000건 이상이어도 전체 행을 대화에 싣지 말고 도구의
상위 결과와 건수만 사용한다. `상태=완료`, 기본적으로 `보류`를 제외한다. 모든 완화 조건 건수도 같은 행 순회에서 계산한다.

## 추가·수정·완료

쓰기 전에 `inspect` 또는 `hash`로 대상과 해시를 얻고, 사용자의 확인 뒤 같은 해시를 `--expected-sha`로 전달한다.

```bash
python3 <tool> inspect --workbook /절대경로/매물장.xlsx --id P001
python3 <tool> add --workbook /절대경로/매물장.xlsx \
  --record-json '<객체 또는 객체 배열>' --expected-sha <sha>
python3 <tool> update --workbook /절대경로/매물장.xlsx --id P001 \
  --changes-json '<객체>' --expected-sha <sha>
python3 <tool> complete --workbook /절대경로/매물장.xlsx --id P001 --expected-sha <sha>
python3 <tool> detail-upsert --workbook /절대경로/매물장.xlsx --id P001 \
  --changes-json '<객체>' --expected-sha <sha>
```

도구는 파일 잠금, 확인 시점 해시, `.maemul-backups` 백업, 임시 파일 교체, 사후 검증을 사용한다. Excel에서 파일이
열린 채 외부 변경이 생겨 해시가 달라지면 쓰지 말고 다시 읽어 변경안을 갱신한다.

## 채팅 입력

한 건은 JSON 객체, 여러 건은 객체 배열로 정규화해 `add`에 전달한다. 없는 값은 생략하거나 `null`로 전달한다. 핵심값이
빠진 행은 도구가 `보류`로 저장하고 경고를 돌려준다. 사용자에게는 "저장하지 못함"이 아니라 "저장됨, 확인 필요"로 알린다.

## 대용량 원칙

- Excel 원본 전체를 CSV로 매번 변환하지 않는다.
- 검색 결과는 기본 10건만 반환한다.
- 정렬·필터는 결정적 도구에서 수행하고 원본 행 전체를 모델 컨텍스트에 넣지 않는다.
- 파일 변경은 전체 통합 문서를 저장하므로, 수만 건에서 잦은 동시 수정이 필요하면 Google Sheets가 더 적합하다고 안내한다.
- SQLite는 대안으로 제시하지 않는다.
