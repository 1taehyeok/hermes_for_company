# hermes-ops

Hermes 운영 규칙, 인수인계 문서, 템플릿을 GitHub로 관리하기 위한 저장소입니다.

## 목적
- 업무별 인수인계 문서를 Markdown으로 표준화
- Hermes가 따라야 할 업무 원칙을 문서화
- 반복 업무용 템플릿을 재사용 가능하게 관리
- GitHub 이력으로 변경사항을 추적

## 기본 원칙
- 원본 파일은 삭제하지 않는다.
- 파일 수정이 필요하면 작업폴더를 만들고 복사본 기준으로 작업한다.
- 업무별 문서는 `handover/02_tasks/` 아래에 1파일 1업무 원칙으로 관리한다.
- 공통 규칙은 `handover/00_rules/`에 둔다.
- 템플릿은 `handover/03_templates/`에 둔다.

## 폴더 구조
```text
hermes-ops/
├─ README.md
├─ .gitignore
├─ handover/
│  ├─ 00_rules/
│  ├─ 01_overview/
│  ├─ 02_tasks/
│  ├─ 03_templates/
│  └─ 99_archive/
├─ skills/
├─ soul/
│  └─ SOUL.md
└─ agents/
```

## 문서 작성 규칙
1. 업무 문서는 제목만 보고도 내용을 알 수 있게 작성한다.
2. 처리 절차는 번호 목록으로 작성한다.
3. 체크리스트는 실제 수행 여부를 바로 확인할 수 있게 쓴다.
4. 주의사항에는 실수 사례와 금지사항을 우선 기록한다.
5. 민감정보, 비밀번호, 토큰은 저장하지 않는다.

## 시작 순서
1. `handover/00_rules/`의 규칙 문서 확인
2. `handover/01_overview/role-summary.md`에 역할과 운영 범위 정리
3. `handover/03_templates/`의 템플릿 복사
4. 새 업무를 `handover/02_tasks/task-xxx-업무명.md` 형식으로 작성

## GitHub 연결 전 체크
- `.gitignore` 반영 여부 확인
- 민감정보 포함 여부 확인
- 실운영 경로와 Git 관리 경로를 혼동하지 않기
