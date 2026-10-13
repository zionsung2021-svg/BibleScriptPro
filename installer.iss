#define MyAppName "BibleScript Pro"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "ZionSung"
#define MyAppExeName "run_local.exe"

[Setup]
AppId={{C78D9F2B-4A61-4B9E-B192-9D092A5F8E21}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=C:\BibleScript\dist_installer
OutputBaseFilename=BibleScriptPro_Setup_v1.0.0
Compression=lzma2/ultra64
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
SetupIconFile=C:\BibleScript\favicon.ico

[Tasks]
Name: "desktopicon"; Description: "바탕화면에 바로가기 만들기"; GroupDescription: "추가 설정:"; Flags: checkedonce

[Files]
; 1. PyInstaller 빌드 결과물 전체 복사
Source: "C:\BibleScript\dist\run_local\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; 2. (안전장치) 실제 원본 bible.db를 설치 폴더의 backend 폴더로 확실하게 복사 보장
Source: "C:\BibleScript\backend\bible.db"; DestDir: "{app}\backend"; Flags: ignoreversion
; 3. 바로가기가 참조할 아이콘 파일도 설치 폴더에 함께 복사
Source: "C:\BibleScript\favicon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\favicon.ico"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon; IconFilename: "{app}\favicon.ico"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "설치 완료 후 BibleScript Pro 실행"; Flags: nowait postinstall skipifsilent

[UninstallRun]
; 언인스톨 시작 전 실행 중인 프로세스 강제 종료 (파일 잠김 방지)
Filename: "taskkill.exe"; Parameters: "/F /IM run_local.exe"; Flags: runhidden; RunOnceId: "KillApp"
Filename: "taskkill.exe"; Parameters: "/F /IM POWERPNT.EXE"; Flags: runhidden; RunOnceId: "KillPPT"

[UninstallDelete]
; 실행 중 생성된 잔여 파일 및 run_local.exe 강제 삭제
Type: files; Name: "{app}\run_local.exe"
Type: filesandordirs; Name: "{app}\_internal"
Type: filesandordirs; Name: "{app}\backend"
Type: filesandordirs; Name: "{app}\*"

[Code]
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // 언인스톨 완료 후 남아있는 설치 폴더 강제 완전 삭제
    DelTree(ExpandConstant('{app}'), True, True, True);
  end;
end;