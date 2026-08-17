# maemul-matching

공인중개사가 Google Sheets 또는 로컬 Excel 매물장을 운영하고, 손님 상담 조건으로 매물을 검색·등록·수정·완료 처리할 수 있는 Codex skill입니다.

## 제공 기능

- Google Sheets 또는 로컬 `.xlsx`를 운영 원본으로 사용
- 불완전한 매물정보도 `?`와 `보류` 상태로 보존
- 손님 조건을 필수 조건과 확인 필요 후보로 나누어 검색
- 매물 추가·가격 수정·계약 완료 처리를 안전하게 수행
- 매물장 템플릿과 결정적 로컬 도구·회귀 테스트 제공

SQLite를 운영 원본으로 사용하지 않으며, 기존 매물·프로필·행을 삭제하지 않습니다.

## 설치

Codex skill 디렉터리에 이 저장소를 `maemul-matching`이라는 이름으로 설치합니다.

```bash
git clone https://github.com/jongwoo01/maemul-matching.git ~/.codex/skills/maemul-matching
```

이미 같은 경로에 설치된 경우에는 저장소를 직접 clone으로 덮어쓰지 말고, 기존 폴더를 백업한 뒤 파일을 교체하세요.

## 사용

Codex에서 다음처럼 요청합니다.

```text
$maemul-matching으로 손님 조건에 맞는 매물을 찾아줘.
```

최초 사용 시 Google Sheets와 로컬 Excel 중 운영 원본을 선택합니다. Google Sheets 작업에는 연결된 Google Drive/Sheets 권한이 필요합니다.

## 테스트

외부 연결 없이 로컬 회귀 테스트를 실행할 수 있습니다.

```bash
python3 -m unittest discover -s tests -p 'test_*.py'
```

## 디렉터리

```text
SKILL.md                 # Codex 진입점과 운영 규칙
agents/openai.yaml       # 표시 이름과 기본 프롬프트
references/              # 설정·검색·Google Sheets·Excel 절차
scripts/                 # 결정적 로컬 도구
tests/                   # 로컬 회귀 테스트
assets/                  # 매물장 CSV 템플릿과 샘플
```

개인 프로필과 인증 정보는 이 저장소 밖의 Codex 설정에 저장해야 하며 커밋하지 않습니다.
