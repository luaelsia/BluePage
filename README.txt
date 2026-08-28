# Blue Page

Office 문서를 Microsoft 365 웹 앱으로 열어 주는 Windows 프로그램입니다.

> **안내:** Blue Page는 Microsoft와 공식적으로 제휴·승인·보증된 프로그램이 아닙니다. 개인이 만든 비공식 작업물입니다.

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
- Office 문서를 더블클릭하면 웹 Word, Excel 또는 PowerPoint에서 열립니다.
- 처음 실행할 때 Microsoft 계정으로 로그인합니다.
- 완전히 종료하려면 트레이 아이콘을 우클릭하고 `종료`를 선택합니다.

## 작동 원리

- 실행한 파일을 OneDrive 에 자동으로 업로드 후 실행시킵니다. 
- 로컬 파일과 OneDrive 의 파일이 서로 다른 경우 동기화 설정 창이 팝업됩니다. 

## 제거

다운로드한 폴더의 `uninstall.ps1`을 실행합니다.

## 개발자

직접 빌드하거나 기능을 수정하려면 DEVELOPERS.txt 를 참고하세요.

문의: miniwhalelabs@gmail.com
