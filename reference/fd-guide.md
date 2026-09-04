## 작성원칙
- 입력자료
  - SR 상세 : SR 분석가가 정리한 SR 요약본을 우선 활용, 없는 경우 ITSM SR 상세를 직접 조회하여 작성
  - 비교 분석 결과 : ABAP 소스분석가가 zdown 수정 전 · 후 소스를 비교 분석한 결과(차이점 · 비즈니스적 의미) — Processing Logic(처리로직) 서술의 근거로 사용
  - zdown 원본 html(수정 전 · 후) : prep 폴더에서 확보 (AGENTS.md `## prep 폴더 취득` 참고) —
    Selection Screen · 실행화면의 as-is/to-be 비교 근거로 사용
  - 기존 FD 문서 : prep 폴더에 있는 경우 버전업 대상으로 사용, 없는 경우 `template` 폴더의 FD 템플릿('KT_ERP_BTA_FD_ZSBFMBR0580_[FM] IP 주문실적전표 일괄 취소 프로그램_20260807.doc')
    파일 기준으로 신규 작성 (prep 폴더와 template 폴더는 서로 다른 별개의 파일임 — AGENTS.md
    `## Prep 폴더 취득`·`## 템플릿 파일 취득` 참고)
- zdown 원본 html(수정 전 · 후)이 없으면 System Screen 항목을, 비교 분석 결과가 없으면 Processing 항목을 작성할 수 없으므로 작성 진행 불가
- 파일명은 'KT_ERP_BTA_FD_{프로그램ID}_[{모듈명}] {프로그램명}_{현재일자}.doc'으로 할 것
  ({프로그램ID}·{프로그램명}·{모듈명}의 정의와 도출 규칙은 AGENTS.md `## 프로그램 식별 정보` 참고)
- 하기의 '수정 세부 사항'을 참고하여 문서를 작성할 것.
- 문서 업데이트 시 수정이력 Format을 지켜서 작성할 것. (표 내부는 색상 마킹 대상에서 제외하되 내용은 동일하게 수정함)
- 수정이력 Format
  Start of {현재일자} [{SR NO}]   -> 글자색 빨간색
  {수정사항 입력}                  -> 글자색 검정색
  End of {현재일자} [{SR NO}]     -> 글자색 빨간색
- 날짜는 yyyy.mm.dd 형식으로 작성할 것.

### 필수사항
- 항상 원본 문서의 복사본을 만들어 버전업 후 저장할 것.

### 금지사항
- 원본 문서는 직접 수정 금지.
- 2페이지 표 안의 기존 History(이력)는 삭제 · 수정 금지.
- 수정이력 Format(색상 규칙) 이외의 서식 · 레이아웃 임의 변경 금지.

## 편집 방법(.doc 파일)
- 템플릿은 Microsoft Word 97-2003 문서(.doc, 바이너리 포맷)이므로 python-docx 등 OOXML(.docx) 기반 라이브러리로는 직접 편집 불가
- 실행 환경(Windows + MS Word 설치됨)을 활용하여 pywin32(`win32com.client`)로 Word 애플리케이션을 직접 자동화하여 편집할 것
  - Word 애플리케이션을 백그라운드로 실행(`Application.Visible = False`)한 뒤 대상 파일을 염 (`Documents.Open`)
  - 텍스트 삽입 · 수정은 Range/Find 객체로 수행
  - 수정이력 Format의 색상 지정 시 RGB 값을 직접 계산하지 말고 `win32com.client.constants`의 `wdColorRed` · `wdColorBlack` 상수를 `Font.Color`에 사용할 것 (Word의 색상 값은 RGB가 아닌 BGR 순서라 직접 계산 시 색이 뒤바뀔 수 있음)
  - 저장 시 파일 형식을 반드시 원본과 동일한 Word 97-2003(`wdFormatDocument97`)으로 지정하여 `SaveAs`할 것 — 형식을 지정하지 않으면 다른 포맷(.docx 등)으로 저장될 수 있음
  - 편집이 실패하더라도 Word 프로세스가 잔류하지 않도록 try/finally 구조로 문서 `Close`, 애플리케이션 `Quit`을 반드시 호출할 것

## 수정 세부 사항
### 1페이지
- 작성일자 : 현재일자
- 작성팀 : atlassian-ktds-kms mcp를 활용하여 현재유저가 속해 있는 팀을 찾아서 작성
- 작성자 : 현재유저

### 2페이지 - Document Management
- 표 안의 기존 History는 수정하지 않고, 비어 있는 칸부터 채워나갈 것.
- Version : 0.1씩 증가
- Date : 현재일자
- Author : 현재유저
- Comments : ITSM의 SR 상세 > 요청내용을 분석하여 간략하게 한 줄로 작성.
- 입력할 셀이 없는 경우 하단에 셀을 추가해서 작성.

### Functional Design 개요
- Requested By : ITSM의 SR 요청자
- Req. Dev. Date : 현재일자
- Prepared By : 현재유저
- Developer : 현재유저

### System Screen
- 수정 전후 로직을 비교하여 Selection Screen 또는 실행 화면의 차이가 존재할 경우 수정이력 Format에 맞춰 수정사항을 as-is와 to-be로 작성.
- 화면 수정사항이 있어 캡쳐를 추가해야 하는 경우 에이전트가 직접 화면을 캡쳐할 수 없으므로, {수정사항 입력} 부분에는 사용자가 어떤 화면을 캡쳐해서 삽입해야 하는지 작성.

### Processing
- 표는 입력할 셀이 없는 경우 하단에 셀을 추가해서 작성.
- Input Parameters : 조회조건이 추가된 경우 작성.
- Output Parameters : 조회화면의 필드가 추가된 경우 작성.
- Processing Logic(처리로직) : 수정 전후 로직을 비교하여 비즈니스 상 수정된 내용을 정리하여 작성. 수정이력 Format에 맞춰 작성할 것.

### New Table(CBO)
- 신규로 생성한 테이블이 있는 경우 기존 템플릿의 표 양식을 참고하여 작성. 수정이력 Format에 맞춰 작성할 것.

## 완료조건
- 파일명 규칙('KT_ERP_BTA_FD_{프로그램ID}_[{모듈명}] {프로그램명}_{현재일자}.doc')을 준수한 산출물 파일이 `result/{SR NO}/` 폴더에 저장됨 (AGENTS.md `## 산출물 저장 위치` 참고)
- 1페이지(작성일자 · 작성팀 · 작성자), 2페이지 Document Management 표(Version · Date · Author · Comments), Functional Design 개요(Requested By · Req. Dev. Date · Prepared By · Developer)가 빠짐없이 채워짐
- 수정 전후 차이가 있는 System Screen · Processing 항목이 수정이력 Format(색상 규칙)에 맞춰 작성됨
- 신규 테이블이 있는 경우 New Table(CBO) 항목이 수정이력 Format(색상 규칙)에 맞춰 작성됨  
  — 위 조건을 실제 생성된 파일에서 확인 (AGENTS.md `정직한 보고 규칙` 준수)
