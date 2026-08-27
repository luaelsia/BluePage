# Blue Page

Blue Page는 Windows의 Office 문서를 Microsoft 365 웹 앱으로 열고 OneDrive와 동기화하는 문서 런처입니다.

- Word: `.docx`, `.doc`
- Excel: `.xlsx`, `.xls`, `.xlsm`
- PowerPoint: `.pptx`, `.ppt`

## 다운로드 및 설치

1. [최신 Release](https://github.com/luaelsia/BluePage/releases/latest)에서 `BluePage-v1.0.0-win-x64.zip`을 다운로드합니다.
2. ZIP 파일의 압축을 풉니다.
3. 압축을 푼 폴더에서 `install.ps1`을 실행합니다.
4. Office 문서를 우클릭하고 `연결 프로그램 → 다른 앱 선택 → Blue Page → 항상 이 앱 사용`을 선택합니다.

설치 스크립트 실행이 차단되면 PowerShell에서 다음 명령을 사용하세요.

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

## 사용

- Office 문서를 더블클릭하면 Microsoft 365 웹 앱에서 열립니다.
- 처음 사용할 때 Microsoft 계정 로그인이 필요할 수 있습니다.
- 설정 창에서 파일 연결, 자동 실행, 동기화 주기와 테마를 변경할 수 있습니다.
- 닫기 버튼을 누르면 트레이로 이동합니다. 완전히 종료하려면 트레이 메뉴에서 `종료`를 선택하세요.

## 제거

다운로드한 폴더에서 `uninstall.ps1`을 실행합니다.

```powershell
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
```

## 개발

소스 코드는 MIT License로 공개되어 있습니다. 개선 사항은 Issue 또는 Pull Request로 보내 주세요. 자세한 내용은 `CONTRIBUTING.txt`를 참고하세요.

```powershell
cd src/BluePage
dotnet build -c Release
```

## 개인정보

- Microsoft 계정 비밀번호를 저장하지 않습니다.
- 개인 PC 모드의 로그인 토큰은 Windows 계정으로 암호화해 로컬에 저장합니다.
- 공용 PC에서는 설정 화면의 `공용 PC` 옵션을 사용하세요.

## 문의

MiniWhaleLabs · miniwhalelabs@gmail.com
