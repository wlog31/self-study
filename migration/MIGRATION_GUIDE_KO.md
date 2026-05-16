# GitHub Pages + Apps Script + Google Sheets 전환 가이드

이 가이드는 "월별 스프레드시트 파일 + 일자별 시트" 구조로 기록을 저장하는 최소 운영 버전입니다.

## 1. 이번에 추가된 파일

- `migration/gh-pages/index.html`: 정적 프론트
- `migration/gh-pages/app.js`: API 호출 로직
- `migration/gh-pages/styles.css`: 스타일
- `migration/apps-script/Code.gs`: 백엔드 API
- `migration/apps-script/appsscript.json`: Apps Script 설정

## 2. 저장 구조

- 월이 바뀌면 새 Google Spreadsheet 파일 생성: `SELF_STUDY_LOG_YYYY-MM`
- 각 날짜 첫 기록 시 일자 시트 생성: `MM-dd` (예: `05-16`)
- 월-파일 매핑 인덱스 파일 자동 생성: `SELF_STUDY_MONTH_INDEX`

## 3. Apps Script 설정

1. Google Apps Script 새 프로젝트 생성
2. `migration/apps-script/Code.gs` 내용 붙여넣기
3. 프로젝트 설정 시간대 `Asia/Seoul`
4. `Deploy -> New deployment -> Web app`
5. 실행 권한:
   - "Execute as": `User accessing the web app` 권장
   - "Who has access": 운영 정책에 맞게 설정
6. Web App URL 복사

### Script Properties (권장)

`Project Settings -> Script properties`에 아래 키를 추가:

- `WRITE_TOKEN`: 쓰기/조회 보호용 토큰(임의 긴 문자열)
- `ROOT_FOLDER_ID`: 월별 파일을 모을 Google Drive 폴더 ID (선택)

## 4. GitHub Pages 설정

1. 레포에 `migration/gh-pages/*`를 배포 대상 경로로 사용
2. GitHub Pages 활성화
3. 페이지 접속 후:
   - `Apps Script Web App URL` 입력
   - `Write Token` 입력 (Script Property와 동일)
   - `Save Config`
   - `Health Check` 성공 확인

## 5. API 동작

- `GET ?action=health`
- `POST action=appendLog` + `payload(JSON)`
- `GET ?action=getLogs&from=YYYY-MM-DD&to=YYYY-MM-DD`
- `GET ?action=getMonthFile&date=YYYY-MM-DD`

## 6. "누가 기록했는지" 추적

- 서버는 `Session.getActiveUser().getEmail()` 값을 `actor_email`로 저장 시도
- 단, 배포 권한/계정 정책에 따라 이메일이 빈값일 수 있음
- 이 경우 `actor_name`(클라이언트 전송값)만 남으므로 신뢰도 낮음

권장:

- Google Workspace 계정 기반으로 Web App 접근 제어
- 혹은 프론트에 별도 로그인(OAuth) 붙인 뒤 토큰 검증 로직 추가

## 7. 추가로 꼭 해야 할 일 (운영 전)

1. 접근 제어
   - 공개 URL 유출 시 임의 쓰기 가능하므로 `WRITE_TOKEN` 필수
   - 가능하면 Web App 접근 범위를 조직 내부로 제한

2. 입력 검증 강화
   - `student_id`, `subject`, `duration` 허용 규칙 명확화
   - 필요 시 금칙어/길이/범위 제한 추가

3. 장애 대비
   - `MONTH_INDEX` 백업
   - 잘못된 월 파일 삭제/이동 시 복구 절차 문서화

4. 기존 Flask 기능 이관 계획
   - 현재 코드는 "학습기록 저장/조회" 중심 최소 버전
   - 출결/좌석/권한관리/승인 플로우는 별도 API 설계 후 단계 이관 필요

5. 테스트 체크리스트
   - 월 경계(예: 2026-05-31 -> 2026-06-01) 자동 생성 확인
   - 동일 날짜 다중 저장
   - 조회 기간 제한(최대 62일) 동작 확인
   - 모바일 브라우저 입력/조회 확인
