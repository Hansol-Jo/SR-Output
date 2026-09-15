## 작성원칙
- 입력자료
  - SR 상세 : SR 분석가가 정리한 SR 요약본을 우선 활용, 없는 경우 ITSM SR 상세를 직접 조회하여 작성
  - 비교 분석 결과 : ABAP 소스분석가가 zdown 수정 전 · 후 소스를 비교 분석한 결과(차이점 · 비즈니스적 의미) — Processing Logic(처리로직) 서술의 근거로 사용
  - zdown 원본 html(수정 전 · 후) : prep 폴더에서 확보 (AGENTS.md `## prep 폴더 취득` 참고) —  Selection Screen · 실행화면의 as-is/to-be 비교 근거로 사용
  - 기존 FD 문서 : prep 폴더에 있는 경우 버전업 대상으로 사용, 없는 경우 `template` 폴더의 FD 템플릿('KT_ERP_BTA_FD_ZSBFMBR0580_[FM] IP 주문실적전표 일괄 취소 프로그램_20260807.doc')  파일 기준으로 신규 작성 
  - prep 폴더와 template 폴더는 서로 다른 별개의 파일임 — AGENTS.md `## prep 폴더 취득`·`## 템플릿 파일 취득` 참고
- zdown 원본 html(수정 전 · 후)이 없으면 System Screen 항목을, 비교 분석 결과가 없으면 Processing 항목을 작성할 수 없으므로 작성 진행 불가
- 파일명은 'KT_ERP_BTA_FD_{프로그램ID}_[{모듈명}] {프로그램명}_{현재일자}'으로 할 것
- 파일확장자는 참고한 파일의 확장자를 따라감. doc또는 docx로 할 것.
- {프로그램ID}·{프로그램명}·{모듈명}·{현재유저}의 정의와 도출 규칙은 AGENTS.md `## 프로그램 식별 정보` 참고
- 하기의 '수정 세부 사항'을 참고하여 문서를 작성할 것.
- 문서 업데이트 시 수정이력 Format을 지켜서 작성할 것. (표 내부는 색상 마킹 대상에서 제외하되 내용은 동일하게 수정함)
- 수정이력 Format
  Start of {현재일자} [{SR NO}]   -> 글자색 빨간색
  {수정사항 입력}                  -> 글자색 검정색
  End of {현재일자} [{SR NO}]     -> 글자색 검정색
- **색상 범위(실제 오류 사례: SRM26090988146)**: 빨간색은 `Start of`/`End of` **마커 줄에만** 적용. 그 사이의 `{수정사항 입력}` 본문(한 줄 제목 + `1)`~`n)` 번호목록)은 **반드시 검정색**이어야 함. 블록 전체를 빨강으로 칠하면 안 됨. 삽입 후 본문 첫 줄의 `Font.ColorIndex`가 검정(1)인지 확인할 것.
- **기존 History 색상 보존**: `Find`로 `Start of`/`End of`를 검색·색상 설정할 때 **SR NO가 다른 기존 이력의 색상까지 건드리지 말 것.** `Find.Text`에 SR NO를 포함해 검색하거나, 검색 후 `Font.Color` 설정 전에 해당 Range의 SR NO가 현재 SR과 일치하는지 확인. 기존 History의 `End of`가 검정으로 바뀌면 안 됨.
- 날짜는 yyyy.mm.dd 형식으로 작성할 것.
- 표를 편집할 때는 윗 줄의 빈칸부터 채워서 작성할 것. 빈칸이 없을 경우에만 행을 추가해서 작성할 것.

### 필수사항
- 항상 원본 문서의 복사본을 만들어 버전업 후 저장할 것.
- 모든 문서는 한글로 작성.

### 금지사항
- 원본 문서는 직접 수정 금지.
- 2페이지 표 안의 기존 History(이력)는 삭제 · 수정 금지.
- 수정이력 Format(색상 규칙) 이외의 서식 · 레이아웃 임의 변경 금지.

## 편집 방법(.doc 파일)
- 템플릿은 Microsoft Word 97-2003 문서(.doc, 바이너리 포맷)이므로 python-docx 등 OOXML(.docx) 기반 라이브러리로는 직접 편집 불가
- **Python은 이 환경에서 사용 불가** — `python`은 Microsoft Store 스텁(exit 9009)이고 `pip` · `py` · `conda`가 없어 pywin32 등 어떤 라이브러리도 설치·실행 불가. Windows PowerShell 5.1 + Word COM(`New-Object -ComObject Word.Application`)만 사용할 것
- 실행 환경(Windows + MS Word 설치됨)을 활용하여 Word COM으로 Word 애플리케이션을 직접 자동화하여 편집할 것
  - Open 전 `DisplayAlerts = 0` · `AutomationSecurity = 3` · `Visible = $true` 설정 — 저장 시 뜨는 대화상자가 COM 자동화에서 응답받지 못해 무한 대기하는 것을 방지
  - 텍스트 삽입 · 수정은 Range/Find 객체로 수행
  - 수정이력 Format의 색상 지정 시 RGB 값을 직접 계산하지 말고 Word 상수(`wdColorRed` = 255, `wdColorBlack` = 0)를 `Font.Color`에 사용할 것 (Word의 색상 값은 RGB 가 아닌 BGR 순서라 직접 계산 시 색이 뒤바뀔 수 있음)
  - 저장 시 파일 형식을 반드시 원본과 동일한 Word 97-2003(`wdFormatDocument97` = 0)으로 지정하여 `SaveAs`할 것 — 형식을 지정하지 않으면 다른 포맷(.docx 등)으로 저장될 수 있음
  - 편집 전에 무수정 복사본을 `.temp`에 SaveAs하는 사전 게이트(60초)를 통과한 경우에만 편집 진행 — 실패 시 편집 시도 금지
  - 편집이 실패하더라도 Word 프로세스가 잔류하지 않도록 try/finally 구조로 문서 `Close`, 애플리케이션 `Quit` + `[GC]::Collect()`를 반드시 호출할 것
- `.ps1` 스크립트는 ASCII 문자만 사용하고 한글 텍스트(파일명 · 내용)는 UTF-8 JSON 파일로 분리할 것 — 한글 주석 포함 시 UTF-8 BOM 문제(mojibake)로 경로 오류 발생
- 대괄호 `[ ]`가 들어간 파일명은 `Copy-Item`/`Move-Item`에 `-LiteralPath`를 사용하고, .NET API는 `[IO.File]` 메서드를 사용할 것
- zdown html 소스는 CP949(EUC-KR) 인코딩 — `[System.Text.Encoding]::GetEncoding(949)`로 디코딩 후 사용할 것
- 산출물 4종 병렬 생성 시 Word COM 스크립트(FD · TD) 동시 실행 금지 — 실행 전 `Get-Process WINWORD` 잔류 0 확인 후 단독 실행

### ⚠️ HTML 파일 읽기 주의
- zdown 소스 html 파일을 읽을 때 한꺼번에 많은 파일을 읽으면 오류 발생
- **fewer batch 방식으로 나누어 읽을 것** (한 번에 1~2 개 파일씩)
- **모든 HTML 파일을 완전히 읽은 후에만 산출물 작성을 시작할 것**
- 파일을 모두 읽지 않고 산출물 작성을 시작하면 차이점을 놓칠 수 있음
- 검증 스크립트의 자동 판정 결과만 믿지 말고 수동 확인 필수

## 수정 세부 사항
### 1페이지
- 표지 작성정보(작성일자 · 작성팀 · 작성자)는 **본문 텍스트가 아니라 표 하단 중앙의 텍스트 상자(Text Box) 안에** 작성함.
  - template FD는 `Text Box 13`을 포함. 본문에 `작성일자 : ...` 형태로 박으면 안 됨(실제 오류 사례: SRM26090988146). 기존 문서 본문에 작성정보가 텍스트로 들어가 있으면 삭제 후 텍스트 상자로 옮길 것.
  - 텍스트 상자 서식: 맑은 고딕, **글자크기 14**, 가운데 정렬, 3단락(작성일 / 작성팀 / 작성자).
- 작성일자, 작성팀, 작성자 키워드가 기존 문서에 없을 경우 작성하지 않고 2페이지로 넘어감.
- 작성일자 : `{현재일자}`만 기입 — **`작성일자` 라벨 텍스트를 넣지 말 것**(날짜만)
- 작성팀 : `작성팀 : {팀명}` — atlassian-ktds-kms mcp를 활용하여 {현재유저}가 속해 있는 팀을 찾아서 작성 (AGENTS.md `## 프로그램 식별 정보` 참고)
- 작성자 : `작성자 : {현재유저}` (AGENTS.md `## 프로그램 식별 정보` 참고)
- 작성예시 (텍스트 상자 3단락)
  2026.09.08
  작성팀 : 재무DX서비스팀
  작성자 : 조한솔

### 2페이지 - Document Management
- 표 안의 기존 History는 수정하지 않고, 비어 있는 칸부터 채워나갈 것.
- 수정이력 Format은 적용하지 않음
- Version : 0.1씩 증가
- Date : {현재 일자} — 기존 값이 있어도 무조건 덮어쓸 것
- Author : {현재유저} — 기존 값이 있어도 무조건 덮어쓸 것
- Comments : ITSM의 SR 상세 > 요청내용을 분석하여 10~15자 내외로 작성.
- **절차**:
  1. 기존 표의 빈 칸 확인
  2. 빈 칸이 있으면 거기부터 채움
  3. 빈 칸이 없을 때만 하단에 셀 추가

### Functional Design 개요
- 명칭 칸을 덮어쓰지 말고, 오른쪽 값 칸에 입력
- Requested By : ITSM의 SR 요청자
- Req. Dev. Date : {현재일자} — 기존 값이 있어도 무조건 덮어쓸 것
- Prepared By : {현재유저} — 기존 값이 있어도 무조건 덮어쓸 것
- Developer : {현재유저} — 기존 값이 있어도 무조건 덮어쓸 것

### System Screen
- 수정 전후 로직을 비교하여 Selection Screen 또는 실행 화면의 차이가 존재할 경우 수정이력 Format에 맞춰 수정사항 작성.
- 기존 값이 있는 경우 뒤에 이어서 작성할 것.
- 화면 수정사항이 있어 캡쳐를 추가해야 하는 경우 에이전트가 직접 화면을 캡쳐할 수 없으므로, {수정사항 입력} 부분에는 사용자가 어떤 화면을 캡쳐해서 삽입해야 하는지 작성.
- 예시
  (기존)
    3. System Screen
      정산 내역 조회
      <화면캡쳐>
      주문 실적전표 취소
      <화면캡쳐>

  (수정사항 추가)
    3. System Screen
      정산 내역 조회
      <화면캡쳐>
      주문 실적전표 취소
      <화면캡쳐>
      Start of 2026.07.24 [SRM26072437047]
      주문내역조회 기능 추가
      (변경 후 selection screen을 캡쳐한다)
      End of 2026.07.24 [SRM26072437047]


### Processing
- 표는 입력할 셀이 없는 경우 하단에 셀을 추가해서 작성.
- 4.2 Input Parameters : Selection Screen의 조회조건. 조회조건이 추가된 경우 작성.
- 4.3 Output Parameters : 조회조건 실행 화면. 실행 화면의 필드가 추가된 경우 작성.
- 4.5 Processing Logic(처리로직) : 
  - 기존 값이 있는 경우 뒤에 이어서 작성할 것.
  - 수정 전후 로직을 비교하여 비즈니스 상 수정된 내용을 정리하여 작성. 수정이력 Format에 맞춰 작성할 것.
    1. prep/{SR NO}/{프로그램ID}_before/ 폴더의 HTML 파일 읽기
    2. prep/{SR NO}/{프로그램ID}_after/ 폴더의 HTML 파일 읽기
    3. 양 폴더의 차이점 분석 (비즈니스 관점)
    4. 수정이력 Format에 맞춰 작성
  - **서술 형식**: `[What]`/`[How]` 태그 사용 금지. 번호 목록 앞에 한 줄 제목을 두어 핵심 수정사항을 요약하고, 그 아래 `1)` `2)` `3)` 순으로 처리 절차를 서술. 명사형 종결(`~함`/`~됨`/`~임`) 사용.
  - **기존 History 색상 보존**: `Find`로 `Start of`/`End of`를 검색할 때 **기존 이력(SR NO가 다른 것)의 색상까지 건드리지 말 것.** `Find.Text`에 SR NO를 포함해 검색하거나, 검색 후 `Font.Color`를 설정하기 전에 해당 Range의 SR NO가 현재 SR과 일치하는지 확인 후 설정. 기존 History의 `End of`가 검정으로 바뀌면 안 됨.
  - 예시
    (기존)
     4.5.1 출력 대상 select
      1) 입력 받은 재무영역코드 = P_KOKRS / 회계연도 = P_GJAHR / 취소 월 = P_MONTH
      2) ztsbfmb0505 = ztsbfmb0501에 존재하는 같은 FM 전표 발췌

    (수정사항 추가)
     4.5.1 출력 대상 select
      1) 입력 받은 재무영역코드 = P_KOKRS / 회계연도 = P_GJAHR / 취소 월 = P_MONTH
      2) ztsbfmb0505 = ztsbfmb0501에 존재하는 같은 FM 전표 발췌
      Start of 2026.07.24 [SRM26072437047]
      3) ztsbfmb0504~5-settl_seq(정산연동회차)가 Max인 값 대상
      4) 정산내역 및 실적전표 취소 조회 시 ztsbfmb0508~9(정산처리내역)에 존재하지 않고, ztsbfmb0504~5-stl_canc_flag(정산취소연동 여부) = ''인 데이터만 출력
      5) 실적전표 취소 역분개 조회 시 ztsbfmb0508~9(정산처리내역)에 존재하는 데이터만 출력
      End of 2026.07.24 [SRM26072437047]
      Start of 2026.09.10 [SRM26090988146]
      정산완료 건 취소 전표 미발행 및 부분정산 건 잔액분만 발행
      1) ztsbfmb0501 조회 시 fmstblg = '' 및 finish = '' 조건 적용 — 전표 미발행·미정산 주문만 조회
      2) ztsbfmb0505와 ztsbfmb0509를 zipno, ipitm, perio, settl_seq로 INNER JOIN — fmbelnr_c_r = space 및 settl_seq = MAX(settl_seq) 조건의 최신 정산내역 조회
      3) 정산금액(stl_sales_amt)·정산수량(stl_qty)을 zipno, ipitm 기준으로 COLLECT 합산
      4) 전체취소(p_indi = 0): 주문금액·수량에서 정산금액·수량 차감 / 부분취소(p_indi = 2): 차감 후 잔액에서 부분취소금액 재차감
      5) 차감 후 금액 > 0인 건만 전표 대상 추가 — 정산완료 건 미발행, 부분정산 건 잔액분만 발행
      End of 2026.09.10 [SRM26090988146]

### New Table(CBO)
- 신규로 생성한 테이블이 있는 경우 기존 템플릿의 표 양식을 참고하여 작성. 수정이력 Format에 맞춰 작성할 것.

## 완료조건
- 파일명 규칙('KT_ERP_BTA_FD_{프로그램ID}_[{모듈명}] {프로그램명}_{현재일자}')을 준수한 산출물 파일이 `result/{SR NO}/` 폴더에 저장됨 (AGENTS.md `## 산출물 저장 위치` 참고)
- 1페이지(작성일자 · 작성팀 · 작성자), 2페이지 Document Management 표(Version · Date · Author · Comments), Functional Design 개요(Requested By · Req. Dev. Date · Prepared By · Developer)가 빠짐없이 채워짐
- 수정 전후 차이가 있는 System Screen · Processing 항목이 수정이력 Format(색상 규칙)에 맞춰 작성됨
- 신규 테이블이 있는 경우 New Table(CBO) 항목이 수정이력 Format(색상 규칙)에 맞춰 작성됨  
  — 위 조건을 실제 생성된 파일에서 확인 (AGENTS.md `정직한 보고 규칙` 준수)
