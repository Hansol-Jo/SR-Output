## 작성원칙
- 입력자료
  - SR 상세 : SR 분석가가 정리한 SR 요약본을 우선 활용, 없는 경우 ITSM SR 상세를 직접 조회하여 작성
  - 비교 분석 결과 : ABAP 소스분석가가 zdown 수정 전 · 후 소스를 비교 분석한 결과(차이점 · 비즈니스적 의미) — Design Strategy 서술의 근거로 사용
  - zdown 원본 html(수정 전 · 후) : Prep 폴더에서 확보 (AGENTS.md `## Prep 폴더 취득` 참고) —
    Selection Screen · 실행화면의 as-is/to-be 비교 근거로 사용. 이 중 수정 후 소스는 Program Source에
    그대로 붙여넣는 원본으로도 사용
  - 기존 TD 문서 : Prep 폴더에 있는 경우 버전업 대상으로 사용, 없는 경우 `template` 폴더의 TD 템플릿
    파일 기준으로 신규 작성 (Prep 폴더와 template 폴더는 서로 다른 별개의 파일임 — AGENTS.md
    `## Prep 폴더 취득`·`## 템플릿 파일 취득` 참고)
- zdown 원본 html(수정 전 · 후)이 없으면 Design Strategy 항목을 작성할 수 없고, zdown 수정 후 소스가 없으면 Program Source 항목을 작성할 수 없으므로 작성 진행 불가
- 파일명은 'KT_ERP_BTA_TD_{프로그램ID}_[{모듈명}] {프로그램명}_{현재일자}.doc'으로 할 것
  ({프로그램ID}·{프로그램명}·{모듈명}의 정의와 도출 규칙은 AGENTS.md `## 프로그램 식별 정보` 참고)
- 하기의 '수정 세부 사항'을 참고하여 문서를 작성할 것.
- 문서 업데이트 시 수정이력 Format을 지켜서 작성할 것.
  - 표 내부는 색상 마킹 대상에서 제외하되 내용은 동일하게 수정함. Program Source는 이 색상 기반 수정이력 Format(빨강/검정)을 적용하지 않지만, 별도의 구문 강조 색상(회색 · 파란색 · 검정색)은 적용함 — 자세한 사항은 아래 'Program Source' 항목 참고
- 수정이력 Format (1페이지 ~ Program Object까지 적용, Program Source는 제외)
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
- Program Source에는 zdown이 추출한 수정 후 소스(html에서 추출한 텍스트)를 그대로 붙여 넣음. 붙여 넣은 소스 코드는 수정 · 삭제 · 임의 추가 금지 — 원문 그대로 복사

## 편집 방법(.doc 파일)
- 템플릿은 Microsoft Word 97-2003 문서(.doc, 바이너리 포맷)이므로 python-docx 등 OOXML(.docx) 기반 라이브러리로는 직접 편집 불가
- Program Source의 ABAP 토큰 분류에는 `Pygments` 라이브러리가 필요함 (설치되어 있지 않으면 `pip install pygments`로 설치)
- 실행 환경(Windows + MS Word 설치됨)을 활용하여 pywin32(`win32com.client`)로 Word 애플리케이션을 직접 자동화하여 편집할 것
  - Word 애플리케이션을 백그라운드로 실행(`Application.Visible = False`)한 뒤 대상 파일을 염 (`Documents.Open`)
  - 텍스트 삽입 · 수정은 Range/Find 객체로 수행
  - 수정이력 Format(1페이지 ~ Program Object)의 색상 지정 시 RGB 값을 직접 계산하지 말고 `win32com.client.constants`의 `wdColorRed` · `wdColorBlack` 상수를 `Font.Color`에 사용할 것
  - Program Source의 구문 색상(주석 · 키워드 · 일반 코드)은 아래 'Program Source' 항목에서 정한 고정 RGB 값(회색 · 파란색 · 검정색)을 사용하며, 이 값들도 RGB → BGR 변환 후 `Font.Color`에 지정할 것 — 변환 공식은 `bgr = (blue << 16) | (green << 8) | red` (Word의 `Font.Color`는 0x00BBGGRR 형태의 정수이므로, RGB 순서를 그대로 넣으면 색이 뒤바뀜)
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
- Comments : ITSM의 'SR 상세 > 요청내용' 을 분석하여 간략하게 한 줄로 작성.
- 입력할 셀이 없는 경우 하단에 셀을 추가해서 작성.

### Technical Design 개요
- Developer : 현재유저
- Req. Dev. Date : 현재일자

### Design Strategy
- 수정 전후 로직을 비교하여 Selection Screen 또는 실행 화면의 차이가 존재할 경우 수정이력 Format에 맞춰 수정사항을 as-is와 to-be로 작성.
- 화면 수정사항이 있어 캡쳐를 추가해야 하는 경우 에이전트가 직접 화면을 캡쳐할 수 없으므로, {수정사항 입력} 부분에는 사용자가 어떤 화면을 캡쳐해서 삽입해야 하는지 작성.

### Program Object
- 수정 전후를 비교하여 대표적으로 사용하는 Object가 추가되었을 경우 빈칸부터 추가 기재.
- 대표적이라 함은 화면에 출력되는 주요 데이터의 원천임. 단순 Text만 가져오는 Transport Table 등은 해당하지 않음.

### Program Source
- 기존 TD 문서를 버전업하는 경우 이 섹션에는 직전 버전의 소스가 이미 붙여넣어져 있음 — 1페이지~Program
  Object까지 적용되는 "기존 내용 보존 + 새 내용 추가"(수정이력 Format) 방식과 달리, Program Source는
  기존에 있던 소스 전체를 삭제하고 이번 SR의 수정 후 소스로 완전히 교체함. 기존 소스 뒤에 이어붙이거나
  기존 소스를 남겨둔 채 추가하지 않음
- 프로그램의 Source 전체를 수정 · 삭제 없이 원문 그대로 붙여넣는 파트임. 로직을 비교하거나 변경 지점을 판단 · 마킹하는 작업이 아니라 zdown 수정 후 html의 소스를 그대로 복사하는 작업임
- Word 문서에는 SAP GUI ABAP 에디터와 동일한 구문 강조 색상을 적용함 : 주석 = 회색(RGB 128,128,128), 키워드 = 파란색(RGB 0,0,255), 그 외 일반 코드 = 검정색(RGB 0,0,0)
  - zdown html 자체의 색상(주석 = 파란색, 로직 = 검정색 2색)은 이 목적에 맞지 않으므로 참고하지 않음
    — html에서는 순수 텍스트(소스 코드)만 추출해서 사용함
  - 주석 · 키워드 판별은 임의로 판단하지 말고 Python `Pygments` 라이브러리의 ABAP 렉서를 사용함  
    (`from pygments.lexers import get_lexer_by_name` → `get_lexer_by_name('abap')`로 토큰화)
  - 토큰화 결과에서 `Token.Comment` 계열은 주석(회색), `Token.Keyword` 계열은 키워드(파란색), 나머지 토큰(변수명 · 리터럴 · 연산자 · 공백 등)은 일반 코드(검정색)로 분류
  - Word Range를 토큰 단위로 나누어 텍스트를 삽입하고, 각 Range의 `Font.Color`를 위 분류에 맞는 RGB 값으로 지정 (RGB → BGR 변환 필요 — 자세한 계산법은 아래 '편집 방법' 참고)
  - Program Source 전체에 고정폭 글꼴(예: Courier New)을 적용하여 원본 SAP 소스의 들여쓰기 · 정렬이 유지되도록 함
- 붙여넣는 순서
  1. 메인 소스 : REPORT 선언부 — SE80에서 프로그램을 더블클릭하면 바로 보이는 최상위 소스
     (헤더 주석 · REPORT문 · INCLUDE 목록 · INITIALIZATION 등 이벤트 블록 포함) 전체
  2. 메인 소스에 나열된 INCLUDE 순서 그대로, 각 INCLUDE의 전체 내용을 이어서 붙여넣기  
     (예 : zsbfmbr0580_top → zsbfmbr0580_cls → zsbfmbr0580_scr → zsbfmbr0580_o01 → zsbfmbr0580_i01 → zsbfmbr0580_f01 순으로 전체를 그대로 붙여넣음)

## 완료조건
- 파일명 규칙('KT_ERP_BTA_TD_{프로그램ID}_[{모듈명}] {프로그램명}_{현재일자}.doc')을 준수한 산출물 파일이 생성됨
- 1페이지(작성일자 · 작성팀 · 작성자), 2페이지 Document Management 표(Version · Date · Author · Comments), Technical Design 개요(Developer · Req. Dev. Date)가 빠짐없이 채워짐
- 수정 전후 차이가 있는 Design Strategy 항목이 수정이력 Format(색상 규칙)에 맞춰 작성됨
- 대표 Object가 추가된 경우 Program Object 항목이 작성됨
- Program Source에 이번 SR의 수정 후 소스(메인 소스 + 나열된 모든 INCLUDE, 순서대로)만 포함되고,
  버전업 이전에 있던 직전 버전의 소스는 남아있지 않음 (소스 내 기존 주석 · 이력 표는 그대로 유지되고
  별도로 가공되지 않음)
- Program Source의 주석(회색) · 키워드(파란색) · 일반 코드(검정색)가 Pygments ABAP 렉서 기준으로 구분되어 표시됨  
  — 위 조건을 실제 생성된 파일에서 확인 (AGENTS.md `정직한 보고 규칙` 준수)
