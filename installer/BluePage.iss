; Blue Page 설치 파일(Inno Setup 6) 정의.
;
; installer\build-setup.ps1 이 dotnet publish 결과물을 넣고 이 파일을 컴파일한다.
; 직접 컴파일하려면 /DAppVersion=1.0.1 /DPublishDir=..\src\BluePage\publish 를 넘긴다.
;
; 관리자 권한 없이 현재 사용자 폴더(%LOCALAPPDATA%\Programs\BluePage)에 설치한다.
; 파일 연결 등록/해제는 BluePage.exe --register / --unregister 가 HKCU에만 쓴다.

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#ifndef PublishDir
  #define PublishDir "..\src\BluePage\publish"
#endif
#ifndef OutputDir
  #define OutputDir "..\release"
#endif

#define AppName "Blue Page"
#define AppPublisher "MiniWhaleLabs"
#define AppUrl "https://github.com/luaelsia/BluePage"
#define AppExeName "BluePage.exe"
#define UserDataDirName "Microsoft365OfficeWebLauncher"

[Setup]
AppId={{7D0C5B0A-3A0E-4B6B-9C2E-5B1F0E8A2C11}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppUrl}
AppSupportURL={#AppUrl}/issues
AppUpdatesURL={#AppUrl}/releases
VersionInfoVersion={#AppVersion}
DefaultDirName={localappdata}\Programs\BluePage
DisableDirPage=yes
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir={#OutputDir}
OutputBaseFilename=BluePage-Setup-v{#AppVersion}
SetupIconFile=..\resources\BluePage.ico
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
ShowLanguageDialog=no

[Languages]
Name: "korean"; MessagesFile: "compiler:Languages\Korean.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[CustomMessages]
korean.LaunchAfterInstall=%1 실행
korean.CreateDesktopIcon=바탕 화면에 바로 가기 만들기
korean.RemoveUserData=설정, 로그인 정보, 백업 파일도 함께 삭제할까요?%n%n"아니요"를 누르면 프로그램만 제거하고 다음 폴더는 남깁니다.%n%1
english.LaunchAfterInstall=Launch %1
english.CreateDesktopIcon=Create a desktop shortcut
english.RemoveUserData=Also delete settings, sign-in data and backups?%n%nChoose No to remove only the program and keep this folder:%n%1

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "{#PublishDir}\*"; DestDir: "{app}"; Excludes: "*.pdb"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
; 파일 연결 후보 등록(HKCU). 설치 화면 뒤에서 조용히 끝난다.
Filename: "{app}\{#AppExeName}"; Parameters: "--register"; Flags: runhidden waituntilterminated
; 설치 마지막 화면의 "Blue Page 실행" 체크 항목. 무인 설치(/SILENT)에서는 실행하지 않는다.
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchAfterInstall,{#AppName}}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{app}\{#AppExeName}"; Parameters: "--unregister"; Flags: runhidden waituntilterminated; RunOnceId: "Unregister"

[Code]
// 트레이에 상주하는 이전 버전이 파일을 잠그고 있으면 복사가 실패하므로 먼저 끝낸다.
procedure KillRunningApp();
var
  ResultCode: Integer;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM {#AppExeName}', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(500);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  KillRunningApp();
  Result := '';
end;

function InitializeUninstall(): Boolean;
begin
  KillRunningApp();
  Result := True;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  UserDataDir: String;
begin
  if CurUninstallStep <> usPostUninstall then
    exit;

  UserDataDir := ExpandConstant('{localappdata}\{#UserDataDirName}');
  if not DirExists(UserDataDir) then
    exit;

  // 무인 제거(/SILENT)에서는 묻지 않고 사용자 데이터를 남긴다.
  if UninstallSilent() then
    exit;

  if MsgBox(FmtMessage(CustomMessage('RemoveUserData'), [UserDataDir]), mbConfirmation, MB_YESNO or MB_DEFBUTTON2) = IDYES then
    DelTree(UserDataDir, True, True, True);
end;
