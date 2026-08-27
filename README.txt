# Blue Page

Blue Page는 Windows에서 Office 문서를 Microsoft 365 웹 앱으로 열어 주는 간단한 문서 런처입니다.

지원 형식:

- Word: `.docx`, `.doc`
- Excel: `.xlsx`, `.xls`, `.xlsm`
- PowerPoint: `.pptx`, `.ppt`

## 설치

Windows 10/11 x64와 .NET 8 SDK 이상이 필요합니다.

설치 파일 : installer/install.ps1

설치가 끝나면 탐색기에서 Office 문서를 우클릭하고 다음을 선택합니다.
`연결 프로그램 → 다른 앱 선택 → Blue Page → 항상 이 앱 사용`

## 사용

- Office 문서를 더블클릭하면 Microsoft 365 웹 앱에서 열립니다.
- 처음 사용할 때 Microsoft 계정 로그인이 필요할 수 있습니다.
- Blue Page 창에서 파일 연결, 자동 실행, 동기화와 테마를 설정할 수 있습니다.
- 프로그램은 닫기 버튼을 눌러도 트레이에서 계속 실행됩니다. 완전히 종료하려면 트레이 메뉴에서 `종료`를 선택하세요.

## 제거

제거 파일 : installer/uninstall.ps1

## 개인정보

- Microsoft 계정 비밀번호를 저장하지 않습니다.
- 개인 PC 모드의 로그인 토큰은 Windows 계정으로 암호화되어 로컬에 저장됩니다.
- 공용 PC에서는 설정 화면의 `공용 PC` 옵션을 사용하세요.

## 문의

MiniWhaleLabs · [miniwhalelabs@gmail.com]
