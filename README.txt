# Blue Page

Office 문서를 Microsoft 365 또는 Google Workspace 웹 앱으로 열어 주는 Windows 프로그램입니다.

> **안내:** Blue Page는 Microsoft 또는 Google과 공식적으로 제휴·승인·보증된 프로그램이 아닙니다. 개인이 만든 비공식 작업물입니다.

## 설치 방법

[최신 설치 파일 다운로드]
https://github.com/luaelsia/BluePage/releases/latest/download/BluePage-v1.0.0-win-x64.zip

1. 다운로드한 ZIP의 압축을 풉니다.
2. `install.ps1`을 실행합니다.
3. Office 문서를 우클릭하고 `연결 프로그램` → `다른 앱 선택`에서 `Blue Page`를 선택합니다.
   - 앱 목록의 정렬 순서나 Windows의 목록 갱신 상태에 따라 `Blue Page`가 바로 보이지 않을 수 있습니다.
   - 이 경우 목록을 끝까지 확인하거나 `PC에서 앱 선택`을 누른 뒤
     `%LOCALAPPDATA%\Programs\BluePage\BluePage.exe`를 한 번 선택하세요.
   - 이후에는 `Blue Page`가 연결 프로그램 목록에 표시됩니다. 기본 앱으로 사용하려면 `항상`을 선택하세요.

## 사용법

- 설치 시 Blue Page가 Office 문서의 연결 프로그램 후보로 자동 등록됩니다. Windows에서 사용자가 기본 앱 지정을 실행해야 합니다.
- Office 문서를 더블클릭하면 Microsoft 365 또는 Google Workspace를 선택해 열 수 있습니다.
- Microsoft 계정과 Google 계정은 Blue Page 설정 화면에서 각각 로그인합니다.
- `문서 연결` 탭에서 전체 기본 열기 방식과 확장자별 Microsoft/Google 열기 방식을 지정할 수 있습니다.
- 완전히 종료하려면 트레이 아이콘을 우클릭하고 `종료`를 선택합니다.

## 작동 원리

- Microsoft 365를 선택하면 OneDrive 앱 전용 폴더 안의 `BluePage` 폴더에 파일을 업로드합니다.
- Google Workspace를 선택하면 Google Drive의 `BluePage` 폴더에 파일을 업로드합니다.
- 로컬 파일과 선택한 클라우드의 파일이 서로 다른 경우 동기화 설정 창이 팝업됩니다.
- 같은 로컬 파일의 Microsoft/Google 원격 항목 ID를 모두 기억하며, 마지막으로 선택한 서비스만 자동 동기화합니다.

## Google Drive 설정

Google 로그인은 배포자가 발급한 Google OAuth 데스크톱 앱 자격 증명이 필요합니다.

1. Google Cloud Console에서 Google Drive API를 활성화합니다.
2. OAuth 동의 화면을 구성하고 `데스크톱 앱` 유형의 OAuth 클라이언트를 만듭니다.
3. Blue Page를 한 번 실행한 뒤 `%LOCALAPPDATA%\Microsoft365OfficeWebLauncher\config.json`을 엽니다.
4. `googleAuth.clientId`와 `googleAuth.clientSecret`에 발급받은 값을 입력합니다.
5. Blue Page를 다시 시작하고 설정 화면에서 Google 로그인을 누릅니다.

Blue Page는 앱이 생성하거나 사용자가 앱을 통해 연 파일만 접근하는 `drive.file` 권한을 요청합니다.

## 동기화 안전 백업

- 온라인 파일로 로컬 파일을 덮어쓰기 전, 현재 로컬 파일을 자동으로 백업합니다.
- 충돌 상황에서 생성한 온라인 사본도 원본 문서 폴더가 아닌 백업 폴더에 저장됩니다.
- 백업 위치: `%LOCALAPPDATA%\Microsoft365OfficeWebLauncher\Backups`
- Blue Page의 `바로가기` → `백업 폴더 열기`에서도 바로 확인할 수 있습니다.
- 잘못된 동기화가 발생했다면 문서 이름별 하위 폴더에서 날짜가 표시된 백업 파일을 찾아 원래 위치로 복사하세요.
- 백업에 실패하면 Blue Page는 로컬 파일 덮어쓰기를 중단합니다.

## 제거

다운로드한 폴더의 `uninstall.ps1`을 실행합니다.

## 개발자

직접 빌드하거나 기능을 수정하려면 DEVELOPERS.txt 를 참고하세요.

문의: miniwhalelabs@gmail.com
