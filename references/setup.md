# 매물장 준비

최초 설정과 `profile start-new` 이후에만 이 4단계를 쓴다. 기존 프로필의 검색·수정에는 표시하지 않는다.

## 매물장 준비 1/4 — 저장 방식

사용자가 이미 방식을 지정하지 않았다면 한 번만 선택하게 한다.

> 매물장을 어디에 둘까요?
> 1. Google Sheets — 여러 사람·기기에서 함께 사용
> 2. 이 컴퓨터의 Excel — 지정 폴더에 `.xlsx` 한 파일로 보관

SQLite나 CSV를 선택지로 제시하지 않는다. 선택 뒤에는 해당 방식 문서만 읽는다.

### Google Sheets

`google-sheets-workflow.md`로 연결 도구와 로그인 계정을 확인한다. 연결 실패를 빈 시트 생성으로 우회하지 않는다.

### 로컬 Excel

`excel-workflow.md`로 폴더 또는 기존 `.xlsx` 경로를 받는다. 폴더라면 존재·쓰기 가능 여부를 확인하고, 덮어쓰지 않는
`매물장.xlsx` 이름을 정한다. 사용자가 다른 이름을 지정해도 된다.

## 매물장 준비 2/4 — 자료 받기

다음을 모두 정상 입력으로 받는다.

- 기존 Google 시트 링크
- Excel·CSV·PDF 파일 또는 파일이 든 폴더
- 채팅에 붙여넣은 매물 한 건 또는 여러 건
- 자료 없이 빈 매물장부터 시작

파일·폴더·링크·매물 텍스트만 답해도 직전 질문의 답으로 처리한다. 채팅 원문은 사용자가 요청하지 않으면 원본 그대로
저장하지 않고 필드만 추출한다.

원본은 수정하지 않는다. 변환 결과는 새 Google 시트 또는 새 로컬 Excel에 만든다. 기존 표준 매물장을 연결하는 경우만
그 파일/시트를 운영 원본으로 쓴다.

## 매물장 준비 3/4 — 아는 값만 정리

`sheet-schema.md`에 맞춰 다음 원칙으로 정리한다.

- 확인된 값만 기록한다.
- 모르는 값은 `?`로 둔다.
- 거래 유형상 해당 없는 가격은 빈칸으로 둔다.
- 모호한 원문은 `특약·메모`에 보존한다.
- 의미가 여러 갈래이고 검색·계약 판단이 달라질 때만 질문한다.
- `종류·거래·지역` 중 빠진 값이 있어도 기록하고 `상태=보류`로 둔다.
- 여러 건은 한 번에 정규화해 일괄 추가한다.

완전한 자료가 될 때까지 설정을 막지 않는다. 처리 건수, 확인이 필요한 건수, 원본 미변경 여부를 알린다.

## 매물장 준비 완료 — 생성·연결

### Google

`매물`, `매물상세` 탭을 만들거나 검증하고 Google 프로필을 저장한다.

```bash
python3 <tool> profile set --name <name> --access google-sheet \
  --sheet-id <id> --spreadsheet-url <url> --connector <connector> --account <account> \
  --listing-sheet 매물 --detail-sheet 매물상세 --activate
```

### 로컬 Excel

새 파일은 먼저 생성하고 검증한 뒤 프로필을 저장한다.

```bash
python3 <tool> init-workbook --workbook /절대경로/매물장.xlsx
python3 <tool> profile set --name <name> --access local-xlsx \
  --workbook /절대경로/매물장.xlsx --listing-sheet 매물 --detail-sheet 매물상세 --activate
```

완료 응답에는 운영 원본 링크 또는 클릭 가능한 절대 파일 경로, 이관 건수, 보류 건수, 다음 검색 요청 예시 하나를 넣는다.

## 새 매물장과 복귀

```bash
python3 <tool> profile start-new
python3 <tool> profile next-name --base 기본매물장
python3 <tool> profile cancel-new
python3 <tool> profile activate --name <기존이름>
```

`start-new`는 기존 원본과 프로필을 보존한다. 새 준비 중에는 기존 이름에 `--replace`를 쓰지 않는다. 사용자가
"초기화"만 말해 범위가 모호할 때만 “기존 매물장은 두고 새 매물장을 준비할까요?”라고 한 번 묻는다.
