# 로컬 Excel 운영

로컬 Excel에는 Google Drive 플러그인이나 로그인이 필요 없다. 한 `.xlsx`의 `매물`·`매물상세` 탭이 운영 원본이다.

## 연결·생성

- 새 빈 매물장: 존재하고 쓰기 가능한 폴더만 받아 `create-excel --directory <절대경로>`를 실행한다. 파일·프로필 이름
  충돌은 자동 번호로 피한다. 오타 경로를 새 폴더로 자동 생성하지 않는다.
- 기존 파일: `profile set --access local-xlsx`가 파일 존재, 두 탭과 표준 헤더를 검증한다.
- 저장된 파일이 이동·삭제됐으면 새 파일을 만들지 말고 새 절대경로를 받는다.

생성은 파일 생성 → 두 탭 검증 → 프로필 저장·활성화 순서다. 중간 실패 시 이전 활성 프로필은 그대로다.

## 검색

```bash
python3 <tool> search --workbook /절대경로/매물장.xlsx \
  --sheet 매물 --criteria-json '<JSON>' --limit 10
```

읽기 전용으로 최신 파일을 순회하고 상위 결과와 집계만 사용한다. 기본적으로 완료·보류는 제외한다.

## 변경

```bash
python3 <tool> inspect --workbook /절대경로/매물장.xlsx --id P001
python3 <tool> add --workbook /절대경로/매물장.xlsx --record-json '<JSON>' --expected-sha <sha>
python3 <tool> update --workbook /절대경로/매물장.xlsx --id P001 --changes-json '<JSON>' --expected-sha <sha>
python3 <tool> complete --workbook /절대경로/매물장.xlsx --id P001 --expected-sha <sha>
python3 <tool> detail-upsert --workbook /절대경로/매물장.xlsx --id P001 --changes-json '<JSON>' --expected-sha <sha>
```

변경 전 해시와 대상 행을 확인한다. 도구는 잠금, 백업, 해시 재검증, 임시 파일 원자 교체를 수행하며 실제 변경 셀이나
추가 행만 갱신한다. 다른 셀 값·행 서식·표준 열 밖 수식과 사용자 확장 열은 보존한다. 해시가 달라지면 다시 읽고
변경안을 갱신한다.

핵심값이 빠진 추가·수정은 보류 경고를 반환한다. `detail-upsert`는 `매물` 탭에 없는 번호를 거부한다.

계약 완료와 계약일·실제계약금액 기록을 함께 요청받으면 먼저 두 변경을 하나의 작업으로 확인받는다. `complete` 뒤 반환된 새
해시로 `detail-upsert`를 실행하고 두 탭을 확인한다. 실제 계약금액을 매물의 호가로 추정하지 않는다. 상세 기록만 실패하면
완료 처리를 다시 실행하지 말고, 완료 반영과 상세 미반영을 나눠 알린다.

15,000건 이상이어도 원본 전체를 대화에 싣지 않는다. 공유·동시 편집이 잦으면 Google Sheets가 더 적합할 수 있다.
