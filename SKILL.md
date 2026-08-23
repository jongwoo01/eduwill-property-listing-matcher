---
name: maemul-matching
description: >-
  공인중개사의 매물장을 Google Sheets 또는 지정 폴더의 로컬 Excel로 만들거나 연결하고, 파일·폴더·채팅으로 받은
  완전하거나 불완전한 매물정보를 기록하며, 손님 카톡·상담 메모의 조건으로 검색하고 추가·수정·완료 처리한다.
  사용자가 "매물 찾아줘", "손님 조건에 맞는 물건", "매물장 등록", "구글 시트", "로컬 엑셀", "엑셀 폴더",
  "매물 정보를 붙여넣을게", "매물 하나 더 들어왔어", "가격 수정", "계약됐어", "다른 매물장으로"처럼 자기 매물장
  설정·검색·관리를 요청할 때 사용한다. 시세·광고 문구·계약서·법령·세금만 요청한 경우에는 사용하지 않는다.
---

# 매물 매칭

운영 원본은 사용자가 고른 **Google Sheets** 또는 **로컬 Excel(.xlsx)** 하나다. SQLite나 CSV를 저장 방식으로
제안하지 않는다. 중개사는 보통 매물장이 하나이므로 프로필도 하나가 기본이다.

## 매번 먼저 할 일 — 연결 확인

1. `python3 <skill-dir>/scripts/maemul_tool.py profile show`로 활성 프로필을 읽는다.
2. 프로필이 없으면 `references/setup.md`의 1/4부터 진행한다.
3. 프로필이 있으면 저장 방식을 다시 묻지 말고 **실제로 열리는지**만 확인한다.
   - Google: 연결된 Google Drive 도구의 스프레드시트 메타데이터 조회(`get_spreadsheet_metadata`)로 `매물`·`매물상세`
     탭 이름과 행 수를 받는다. 실패하면 `references/user-experience.md`의 오류 절대로 재연결을 안내하고 멈춘다.
   - Excel: `profile show`의 `workbook_exists`가 `false`면 새 파일을 만들지 말고 새 경로를 받는다.
4. 사용자가 이미 저장 방식이나 링크를 말했으면 같은 질문을 반복하지 않는다.

프로필은 `$CODEX_HOME/maemul-matching/profiles.json` 또는 `~/.codex/maemul-matching/profiles.json`에 저장한다.

## 필요한 자료만 읽기

한 요청에서 모든 참고 문서를 읽지 않는다.

| 작업 | 읽을 문서 |
| --- | --- |
| 최초 설정·새 매물장 | `references/setup.md` + 선택한 저장 방식 문서 |
| 검색 | `references/matching-rules.md` + 선택한 저장 방식 문서의 읽기 절차 |
| 추가·수정·완료 | 선택한 저장 방식 문서의 변경 절차 |
| 생성·이관·헤더 오류 | `references/sheet-schema.md` + `references/tool-usage.md` |
| 안내 문구·오류 복구가 필요함 | `references/user-experience.md`의 해당 절만 |

저장 방식 문서는 `references/google-sheets-workflow.md` 또는 `references/excel-workflow.md` 중 하나만 읽는다.

## 의도와 안전 규칙

- 검색은 확인 없이 읽기 전용으로 실행한다.
- 추가·수정·완료는 대상과 변경 전후를 보여주고 확인받은 뒤 쓴다. 이미 명확히 승인된 생성·일괄 이관은 재확인하지 않는다.
- 사용자가 "확인 없이 바로 해줘"라고 해도 **추가 전 번호 중복 검사 → 쓰기 → 재조회 검증** 세 단계는 생략하지
  않는다. 확인을 생략하는 것이지 검증을 생략하는 것이 아니다.
- 기존 매물장·프로필·행을 삭제하지 않는다. "삭제해줘"는 `상태=완료`(거래 종료) 또는 `상태=보류`(잠시 빼둠)로
  바꾸는 것으로 안내하고, 행 자체는 남긴다.
- 다른 매물장으로 바꾸는 요청은 `profile activate`, 새 매물장은 기존 프로필을 남긴 채 **새 이름**으로
  `profile set --activate`다. 기존 이름에 `--replace`를 쓰려면 사용자에게 덮어쓴다는 점을 확인받는다.
- 파일·폴더·시트 링크나 매물 텍스트만 와도 직전 질문의 답으로 보고 계속한다.
- 채팅 입력은 한 건과 여러 건 모두 받는다. 아는 값만 기록하고 모르는 값은 `?`, 해당 없는 가격은 빈칸으로 둔다.
- `종류·거래·지역` 중 하나라도 없으면 저장을 막지 않고 `상태=보류`로 기록한다. 보류는 기본 검색에서 제외한다.
- 사용자의 표현이 여러 값으로 갈려 검색이나 계약에 영향을 줄 때만 질문한다. 나머지는 원문을 메모에 보존한다.

## 검색

1. `references/matching-rules.md`로 자연어를 하드·소프트 조건으로 만든다.
2. 활성 원본의 최신 `매물` 탭만 읽는다. Google은 필요한 범위/청크만, Excel은 로컬 도구의 읽기 전용 검색을 쓴다.
3. `상태=완료`와 기본적으로 `보류`를 제외한다.
4. 하드 조건이 모두 확인된 항목과 값 확인이 필요한 후보를 분리한다.
5. 상위 3건은 근거를 자세히, 나머지는 합계 10건까지 간단히 보여준다. 개인정보는 제외한다.
6. `?`가 있는 후보는 확인 전까지 적합하다고 단정하지 않는다.

## 변경

- Google은 라이브 대상 행 재조회 → 번호 중복 검사(추가) → 변경안 확인 → 범위가 좁은 배치 쓰기 → 재조회 검증 순서다.
- Excel은 `inspect`/`hash` → 변경안 확인 → `--expected-sha` 변경 → 백업과 결과 해시 확인 순서다.
- 소유자·연락처 등 개인정보는 `매물상세`에만 둔다.
- 계약가·계약일을 모르면 호가를 대신 쓰지 않고 `?`로 둔다.

## 참고 자료

- `references/setup.md` — 저장 방식 선택과 4단계 준비, 매물장 전환
- `references/google-sheets-workflow.md` — Google 읽기·쓰기
- `references/excel-workflow.md` — 폴더 기반 Excel 읽기·쓰기
- `references/matching-rules.md` — 조건과 순위
- `references/sheet-schema.md` — 두 탭의 컬럼과 불완전 데이터
- `references/tool-usage.md` — 프로필·Excel 명령
- `references/user-experience.md` — 짧은 안내·복구 문구
- `scripts/maemul_tool.py` — 결정적 검색과 안전한 로컬 변경
