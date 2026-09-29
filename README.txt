# Blue Page

Office 문서를 Microsoft 365 또는 Google Workspace 웹 앱으로 열어 주는 Windows 프로그램입니다.

> Microsoft, Google과 제휴하거나 승인받은 프로그램이 아닌 개인 제작물입니다.

## 설치

1. [최신 릴리스](https://github.com/luaelsia/BluePage/releases/latest)에서 `BluePage-Setup-vX.Y.Z.exe`를 받아 실행합니다. 관리자 권한은 필요 없습니다.
   - SmartScreen 경고가 뜨면 `추가 정보` > `실행`을 누릅니다(코드 서명이 없어서 뜨는 경고).
2. Office 문서를 우클릭하고 `연결 프로그램` > `다른 앱 선택`에서 `Blue Page`를 고른 뒤 `항상`을 누릅니다.

새 버전이 나오면 앱이 알려 주고, 누르면 바로 설치합니다.

## 사용

- 문서를 더블클릭하면 Microsoft 365 또는 Google Workspace로 열립니다. 편집한 내용은 로컬 파일과 자동으로 동기화됩니다.
- 계정 로그인은 `홈`, 열 서비스 지정은 `문서 연결`에서 합니다.
- 완전히 끄려면 트레이 아이콘을 우클릭하고 `종료`를 누릅니다.

## 백업

로컬 파일을 온라인 내용으로 덮어쓰기 전의 파일과 충돌 사본은 `%LOCALAPPDATA%\Microsoft365OfficeWebLauncher\Backups`에 보관합니다. 트레이 메뉴의 `백업 폴더 열기`로 열 수 있습니다.

## 제거

Windows `설정` > `앱` > `설치된 앱`에서 `Blue Page`를 제거합니다.

## 개발

빌드와 Google OAuth 설정은 DEVELOPERS.txt를 참고하세요.

문의: miniwhalelabs@gmail.com
