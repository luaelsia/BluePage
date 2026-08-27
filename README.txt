# Blue Page

**Blue Page**는 MiniWhaleLabs가 만드는 Windows용 문서 런처입니다. 문의: [miniwhalelabs@gmail.com](mailto:miniwhalelabs@gmail.com)

Windows 탐색기에서 Office 문서(.docx/.xlsx/.pptx 등)를 더블클릭하면 로컬 Office 데스크톱 앱 대신 **Microsoft 365 Office Web**(Word/Excel/PowerPoint Online)에서 열어 주는 Windows용 런처입니다.

> 구현 전 조사 결과와 설계 근거는 [`docs/RESEARCH.txt`](docs/RESEARCH.txt), 아키텍처 상세는 [`docs/ARCHITECTURE.txt`](docs/ARCHITECTURE.txt)를 참고하세요.

## 1. 프로그램 개요

- 지원 문서: Word(.docx/.doc), Excel(.xlsx/.xls/.xlsm), PowerPoint(.pptx/.ppt) — `appsettings.json`에 항목 추가만으로 확장 가능
- 구현 언어: **C# / .NET 8**(WinExe, win-x64). 이유는 아래 "왜 C++가 아닌 C#인가" 참고
- Office 데스크톱 앱을 전혀 사용하지 않음
- 로그인 자격 증명(ID/PW)을 앱에 저장하지 않음 — Windows 계정 기반 사일런트 SSO(MSAL + WAM)

### 왜 C++가 아닌 C#인가

Microsoft는 인증(MSAL)과 Microsoft Graph SDK를 .NET/JS/Java/Python/iOS·macOS용으로만 공식 배포하며, **C++용 공식 SDK는 존재하지 않습니다.** 인증·업로드처럼 보안이 중요한 로직을 직접 구현하는 위험을 피하기 위해 C# / .NET 8을 선택했습니다. `dotnet publish`로 런타임 설치가 필요 없는 배포물을 만들 수 있고, Win32 레지스트리·ShellExecute·WebView2 등 Windows 네이티브 API도 .NET에서 동일하게 사용할 수 있습니다.

## 2. 빌드 방법

사전 요구사항: .NET 8 SDK 이상(Windows), Windows 10/11 x64.

```powershell
cd BluePage/src/BluePage
dotnet build -c Release
```

배포용(런타임 설치 불필요한 단일 실행형) 게시:

```powershell
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishTrimmed=false -o publish
```

> **참고**: `PublishSingleFile`을 사용해도 WPF/WindowsForms 렌더러와 MSAL WAM 브로커의 네이티브 구성 요소는 .NET/Windows 플랫폼 제약상 `BluePage.exe` 옆에 함께 배치될 수 있습니다. 따라서 설치할 때는 `publish` 폴더 전체를 복사해야 합니다.

## 3. 설치

```powershell
# publish 폴더가 준비된 상태에서
./installer/install.ps1
```

`%LOCALAPPDATA%\Programs\BluePage`에 설치되고(관리자 권한 불필요), 파일 연결 등록까지 자동 수행합니다. 제거는 `./installer/uninstall.ps1`.

## 4. Entra App Registration 설정 (배포자가 최초 1회만 수행 — 이미 완료됨)

Microsoft Graph API를 호출하려면 Microsoft Entra ID(구 Azure AD)에 애플리케이션을 등록해 Client ID를 발급받아야 합니다. **이 저장소는 이미 등록을 완료해 발급받은 Client ID를 `resources/appsettings.default.json`에 포함해 빌드하므로, 이 실행 파일을 쓰는 사람은 아래 절차를 다시 할 필요가 없습니다.** GUI에도 Client ID 입력란이 없습니다.

이 저장소를 포크하거나 자신만의 Client ID로 다시 빌드하고 싶다면 아래 절차를 따르세요(사용자가 직접 수행해야 하며 자동화할 수 없음):

1. https://entra.microsoft.com 에 로그인
2. **(2024년 6월 이후 정책 변경)** 개인 Microsoft 계정만으로는 "디렉터리 없이" 앱을 등록할 수 없습니다. 화면에 "디렉터리 외부에서 애플리케이션을 만드는 기능은 더 이상 사용되지 않습니다"라는 메시지가 뜬다면 디렉터리(테넌트)가 없는 것이므로, 먼저 아래 중 하나로 디렉터리를 만드세요.
   - **Microsoft 365 개발자 프로그램**(https://developer.microsoft.com/microsoft-365/dev-program) 가입 — 다만 2025년경부터 무료 샌드박스 테넌트 발급 자격이 Visual Studio 구독자 등으로 제한되어, 일반 개인 계정은 "You don't currently qualify" 메시지와 함께 거부될 수 있습니다.
   - 위 방법이 안 되면 **Azure 무료 계정**(https://azure.microsoft.com/free) 가입 — 가입 시 본인 확인용 카드·휴대폰 번호가 필요하지만(실제 청구 없음, 임시 홀드 후 자동 취소) 가입만 하면 **Microsoft Entra ID(Free)** 테넌트가 자동 생성됩니다.
3. 디렉터리가 생긴 계정으로 https://entra.microsoft.com 재로그인 → **ID > 앱 등록(App registrations) > 새 등록(New registration)**
4. 이름: 자유롭게 입력 (예: `Microsoft365OfficeWebLauncher`)
5. 지원되는 계정 유형: **"모든 Entra ID 테넌트 + 개인 Microsoft 계정"** 선택 (회사/학교 계정과 개인 계정 모두 지원하기 위함 — 앱 등록이 이 테넌트에 있어도 실제 로그인은 사용자의 평소 OneDrive 계정으로 하면 됨)
6. 리디렉션 URI는 이 단계에서 비워두고 등록 완료
7. 등록 후 **개요** 페이지에서 **애플리케이션(클라이언트) ID**를 복사
8. **Authentication > 플랫폼 추가 > 모바일 및 데스크톱 애플리케이션**
   - 리디렉션 URI로 `ms-appx-web://microsoft.aad.brokerplugin/<복사한 클라이언트 ID>` 를 직접 입력(WAM 브로커 필수 URI)
   - 하단 **"공용 클라이언트 흐름 허용"을 "예"** 로 설정
9. **API 권한(API permissions) > 권한 추가 > Microsoft Graph > 위임된 권한(Delegated permissions)** 에서 `Files.ReadWrite.AppFolder` 검색 후 추가 (관리자 동의가 필요 없는 일반 권한)
10. 발급받은 Client ID를 `resources/appsettings.default.json`의 `auth.clientId`에 넣고 다시 빌드/게시

- **개인 PC**: `sharedPcMode: false`(기본값) — 최초 1회 로그인 후 암호화된 토큰 캐시로 이후 자동(사일런트) 로그인
- **공용 PC**: GUI의 "공용 PC" 체크박스를 켜면 즉시 저장된 로그인 정보를 삭제하고, 다음 실행부터 토큰을 디스크에 저장하지 않고 매번 로그인 요구(다른 이용자와 계정 분리)

## 5. 사용법

```powershell
BluePage.exe "C:\Docs\Report.docx"     # 문서를 Office Web에서 엽니다
BluePage.exe --register                # 파일 형식을 연결 프로그램 후보로 등록
BluePage.exe --unregister              # 등록 해제
BluePage.exe --sync "C:\Docs\Report.docx"  # 브라우저를 열지 않고 즉시 동기화만 수행
BluePage.exe --sync-all                # 등록된 모든 파일을 일괄 동기화
BluePage.exe --settings                # 상태/설정 창을 엽니다
BluePage.exe --minimized               # 창을 띄우지 않고 바로 트레이로 시작(Windows 자동 시작용)
```

`--register` 실행 후 탐색기에서 문서 우클릭 → **연결 프로그램 → 다른 앱 선택** → *Blue Page* 선택 → **"항상 이 앱 사용"** 체크 (Windows 정책상 앱이 스스로 기본값을 강제 지정할 수 없어 이 1회 확인이 필요합니다. [§6](#6-microsoft-구조상-불가능한-부분) 참고)

GUI(인수 없이 실행)에는 **"Windows 시작 시 자동 실행"** 체크박스가 있습니다(`HKCU\Software\Microsoft\Windows\CurrentVersion\Run`에 등록, 관리자 권한 불필요). 함께 있는 **"트레이로 최소화해서 시작"**을 켜면 부팅 시 창이 뜨지 않고 바로 트레이 아이콘으로 상주합니다. 창을 닫아도(X 버튼) 앱은 종료되지 않고 트레이에 남으며, 완전히 끄려면 트레이 아이콘 우클릭 → 종료를 사용합니다.

**트레이 아이콘은 항상 떠 있고(프로그램이 켜져 있는 동안 계속), 우클릭하면 창을 열지 않고도 거의 모든 기능을 바로 쓸 수 있습니다**: 로그인/로그아웃, 파일 연결 등록/해제, 기본 프로그램으로 설정, 동기화 검토, OneDrive에서 보기, 자동 동기화 주기(프리셋), Windows 자동 실행/트레이 시작/공용 PC 모드 켜고 끄기, 로그 폴더 열기, 종료. 트레이 메뉴와 창의 컨트롤은 같은 상태를 공유하므로 어느 쪽에서 바꿔도 서로 동기화됩니다.

## 6. 동작 원리 / 파일 처리 흐름

```
더블클릭(.docx)
  → BluePage.exe "<경로>"
  → 확장자로 Office 앱 판별 (Word/Excel/PowerPoint)
  → MSAL + WAM으로 Microsoft 365 계정 토큰 획득 (최초 1회 로그인, 이후 사일런트)
  → 로컬↔온라인 상태 비교(4가지 중 하나로 분기):
      · 신규 파일           → OneDrive App Folder에 새로 업로드
      · 변경 없음           → 기존 온라인 사본 그대로 사용
      · Web에서만 수정됨    → 온라인 사본을 로컬로 자동 다운로드(반영) 후 오픈
      · 로컬에서만 수정됨   → 로컬 내용을 온라인으로 재업로드
      · 양쪽 다 수정됨(충돌) → 아래 4가지 중 하나를 사용자에게 물어봄
  → 응답받은 webUrl을 기본 브라우저로 열기 (Process.Start)
  → 각 단계를 %LOCALAPPDATA%\Microsoft365OfficeWebLauncher\logs 에 기록
```

**충돌(양쪽 다 수정됨) 발생 시 선택지** — 창이 하나 뜨며, 기본 선택은 항상 가장 안전한 "사본 생성"입니다.

| 선택지 | 동작 |
|---|---|
| **사본 생성** (기본값) | 온라인 사본을 별도 파일로 저장하고 로컬 원본은 그대로 둠(둘 다 보존) |
| 더 최신 파일로 덮어쓰기 | 로컬/온라인 중 수정 시각이 더 나중인 쪽으로 다른 쪽을 덮어씀 |
| 오프라인(로컬) 파일로 덮어쓰기 | 로컬 파일 내용으로 온라인 사본을 덮어씀 |
| 온라인 파일로 덮어쓰기 | 온라인 사본 내용으로 로컬 파일을 덮어씀 |

자세한 판정 로직은 [`docs/ARCHITECTURE.txt`](docs/ARCHITECTURE.txt)를 참고하세요.

**백그라운드 자동 동기화**: GUI가 떠 있거나 트레이에 상주하는 동안 일정 주기로 위와 같은 방식으로 추적 중인 모든 파일을 자동으로 확인/반영합니다(로그인돼 있고 추적 중인 파일이 있을 때만 동작). Graph의 실시간 변경 알림(webhook)은 공개 HTTPS 엔드포인트가 필요해 로컬 앱에는 쓸 수 없어 폴링 방식을 사용하며, 문서를 다시 열지 않고도 어느 정도 실시간에 가깝게 반영됩니다. 주기는 GUI의 **"자동 동기화 주기(초)"** 입력으로 조절할 수 있고(기본 180초 = 3분, 1~600초 범위, `config.json`의 `backgroundSyncIntervalSeconds`), 변경하면 재시작 없이 바로 적용됩니다.

**GUI "동기화 검토…" 버튼**: 위 자동 동기화(더블클릭 오픈/백그라운드 폴링/`--sync-all` CLI)와 달리, 실행 전에 파일별 감지 상태·마지막 동기화 시각·실행할 동작을 미리 보여주는 검토 창이 뜹니다. 파일마다 "건너뛰기"를 포함한 선택지 중 원하는 동작을 직접 고를 수 있고, `[선택 항목 동기화]`를 눌러야 실제로 반영됩니다 — 원하지 않는 파일이나 방향으로 자동 처리되는 것을 막기 위한 유일한 수동/검토용 진입점입니다.

## 7. Microsoft 구조상 불가능한 부분

1. **로컬 파일을 업로드 없이 여는 것 자체가 불가능합니다.** Office Web은 브라우저 보안 모델상 로컬 파일 시스템에 접근할 수 없고, 반드시 OneDrive/SharePoint 등 온라인 위치의 항목이어야 합니다. Microsoft가 만든 공식 Office Web PWA조차 이 제약을 그대로 갖고 있음을 확인했습니다([`docs/RESEARCH.txt`](docs/RESEARCH.txt) 5절).
2. **앱이 스스로를 "기본 프로그램"으로 강제 지정할 수 없습니다.** Windows 8 이후 정책으로, 사용자가 설정 앱 또는 "연결 프로그램"에서 1회 확인해야 합니다.
3. **실시간(webhook) 자동 동기화는 제공하지 않습니다.** Microsoft Graph의 변경 알림은 공개 HTTPS 엔드포인트가 필요해 순수 로컬 앱에 적용할 수 없으므로, 대신 트레이에 상주하며 주기적으로 확인하는 폴링 방식(§6 "백그라운드 자동 동기화")으로 동작합니다.
4. **비밀번호 자동 입력을 통한 로그인은 구현하지 않았습니다.** 요구사항에는 "미리 설정한 ID/PW로 자동 로그인"이 있었으나, 이는 평문/역가역 자격증명 저장이라는 보안 위험이 있어 채택하지 않고 MSAL + WAM 사일런트 SSO로 대체했습니다(§8 참고).
5. **온라인 사본이 Office Web에서 열려 편집 중인 동안은 로컬 변경사항을 그 위에 덮어쓸 수 없습니다.** OneDrive/SharePoint가 공동 작성 세션 중인 파일에 대한 직접 콘텐츠 교체 요청을 HTTP 423(Locked)으로 거부하기 때문으로, Microsoft 서버 측 정책이라 우회할 수 없습니다. 이 경우 더블클릭으로 열면 반영은 실패하지만 기존 온라인 사본은 그대로 열어 주고, "동기화 검토…"에서는 실패한 파일을 요약해서 알려줍니다. 편집을 마치고(온라인 세션이 끝나고) 나면 다음 시도부터 정상적으로 반영됩니다.

## 8. 공식 지원 기능

- Microsoft Graph API(`driveItem` 업로드/다운로드, `webUrl`) — 공식 문서에 기술된 표준 사용법
- MSAL(Microsoft Authentication Library) + WAM(Web Account Manager) 브로커 — Windows 계정 통합 사일런트 SSO의 공식 권장 방식
- OneDrive App Folder(`Files.ReadWrite.AppFolder`) — 최소 권한 원칙에 따른 공식 지원 스코프
- HKCU 기반 파일 연결 등록(`RegisteredApplications`, `Applications\<exe>`) — Windows에서 공식적으로 문서화된 "연결 프로그램 후보 등록" 방법

## 9. 비공식 우회 방법 (검토했으나 채택하지 않음)

- WOPI 프로토콜로 자체 호스트에서 Office Web 임베드: 기술적으로는 가능하나 Microsoft 365 Cloud Storage Partner Program 승인이 필요해 개인/사내 런처에는 비현실적이라 제외
- WebView2로 `file://` 문서를 Office Web에 직접 로드: Microsoft가 지원하지 않으며 실제로는 다운로드로 대체됨을 확인해 제외
- Edge PWA의 비공식 커맨드라인 플래그로 로컬 파일 오픈 우회: 존재하지 않는 기능으로 확인되어 제외

## 10. 향후 개선 방향

- **WebView2 임베드 UX**: 브라우저 전환 없이 런처 자체 창에서 Office Web을 띄우는 통합 경험
- **Outlook/Visio/OneNote 등 추가 Office 앱 지원**: `appsettings.json`의 `documentTypes`에 항목 추가만으로 확장 가능하도록 이미 설계되어 있어, 실제 앱별 URL 동작 검증만 남아 있음
- **업로드/동기화 진행률 GUI 토스트**: 현재는 로그 파일과 완료 후 메시지 박스로만 안내

## 11. 로그 / 문제 해결

- 로그 위치: `%LOCALAPPDATA%\Microsoft365OfficeWebLauncher\logs\launcher-YYYYMMDD.log` (일자별 롤링, 기본 14일 보관)
- 설정 파일: `%LOCALAPPDATA%\Microsoft365OfficeWebLauncher\config.json`
- 업로드 매핑: `%LOCALAPPDATA%\Microsoft365OfficeWebLauncher\manifest.json`
- 토큰 캐시(개인 PC 모드): `%LOCALAPPDATA%\Microsoft365OfficeWebLauncher\tokencache\msal_token_cache.bin` (Windows DPAPI로 암호화, 비밀번호 자체는 저장되지 않음)

## 12. 프로젝트 구조

```
BluePage/
  src/BluePage/       # C# 소스 (Core/Auth/OneDrive/Config/Registry/Logging)
  include/            # C# 채택으로 현재 비어 있음(향후 네이티브 확장 헤더용)
  resources/          # 최종 앱 아이콘(BluePage.png/.ico), 크기별 아이콘, 기본 설정
  installer/          # install.ps1 / uninstall.ps1
  docs/               # RESEARCH.txt, ARCHITECTURE.txt
  README.txt          # 이 문서
```
