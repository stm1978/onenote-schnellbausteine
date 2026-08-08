#ifndef AppVersion
  #define AppVersion "0.1.0"
#endif
#ifndef FileVersion
  #define FileVersion "0.1.0.0"
#endif

#define AppName "Schnellbausteine für OneNote"
#define AppPublisher "Schnellbausteine"
#define AppIdValue "{{5B1518D1-0620-49BC-A423-0684C3E21917}"
#define AddInClsid "{{CCF3B64A-9580-4D44-BDCE-49BFE95AC510}"
#define AddInProgId "Schnellbausteine.OneNoteAddIn"
#define AddInAssembly "Schnellbausteine.AddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"

[Setup]
AppId={#AppIdValue}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=https://github.com/stm1978/onenote-schnellbausteine
AppSupportURL=https://github.com/stm1978/onenote-schnellbausteine/issues
AppUpdatesURL=https://github.com/stm1978/onenote-schnellbausteine/releases
DefaultDirName={localappdata}\Schnellbausteine\AddIn
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x86compatible x64compatible
CloseApplications=no
CloseApplicationsFilter=ONENOTE.EXE
RestartApplications=no
OutputDir=..\artifacts\setup
OutputBaseFilename=Schnellbausteine-Setup-{#AppVersion}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes
UninstallDisplayName={#AppName}
VersionInfoVersion={#FileVersion}
VersionInfoCompany={#AppPublisher}
VersionInfoDescription=Installer für das OneNote-Add-in Schnellbausteine
VersionInfoProductName={#AppName}
VersionInfoProductVersion={#FileVersion}
VersionInfoCopyright=Copyright (c) 2026
MinVersion=10.0

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Messages]
german.WelcomeLabel1=Willkommen beim Setup von [name]
german.WelcomeLabel2=Dieses Programm installiert das OneNote-Add-in Schnellbausteine nur für Ihr Windows-Benutzerkonto.%n%nDie passende 32- oder 64-Bit-Version wird automatisch anhand Ihrer OneNote-Installation ausgewählt.%n%nOneNote muss während Installation und Deinstallation vollständig geschlossen sein.
german.FinishedHeadingLabel=Installation abgeschlossen
german.FinishedLabelNoIcons=Schnellbausteine wurde erfolgreich für OneNote installiert.%n%nStarten Sie OneNote jetzt neu.

[Files]
Source: "..\artifacts\x86\Release\Schnellbausteine.AddIn.dll"; DestDir: "{app}\x86"; Flags: ignoreversion; Check: InstallX86
Source: "..\artifacts\x86\Release\Schnellbausteine.Core.dll"; DestDir: "{app}\x86"; Flags: ignoreversion; Check: InstallX86
Source: "..\artifacts\x64\Release\Schnellbausteine.AddIn.dll"; DestDir: "{app}\x64"; Flags: ignoreversion; Check: InstallX64
Source: "..\artifacts\x64\Release\Schnellbausteine.Core.dll"; DestDir: "{app}\x64"; Flags: ignoreversion; Check: InstallX64

[InstallDelete]
Type: filesandordirs; Name: "{app}\x86"; Check: InstallX64
Type: filesandordirs; Name: "{app}\x64"; Check: InstallX86

[Registry]
; Remove stale registration in both views before writing the detected architecture.
Root: HKCU32; Subkey: "Software\Microsoft\Office\OneNote\AddIns\{#AddInProgId}"; Flags: deletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddInClsid}"; Flags: deletekey
Root: HKCU32; Subkey: "Software\Classes\AppID\{#AddInClsid}"; Flags: deletekey
Root: HKCU32; Subkey: "Software\Classes\{#AddInProgId}"; Flags: deletekey
Root: HKCU64; Subkey: "Software\Microsoft\Office\OneNote\AddIns\{#AddInProgId}"; Flags: deletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddInClsid}"; Flags: deletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\AppID\{#AddInClsid}"; Flags: deletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\{#AddInProgId}"; Flags: deletekey; Check: IsWin64

; 32-bit OneNote registration.
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddInClsid}"; ValueType: string; ValueData: "Schnellbausteine für OneNote"; Flags: uninsdeletekey; Check: InstallX86
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddInClsid}"; ValueType: string; ValueName: "AppID"; ValueData: "{#AddInClsid}"; Check: InstallX86
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddInClsid}\ProgID"; ValueType: string; ValueData: "{#AddInProgId}"; Check: InstallX86
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddInClsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"; Check: InstallX86
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddInClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: InstallX86
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddInClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Schnellbausteine.AddIn.AddIn"; Check: InstallX86
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddInClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AddInAssembly}"; Check: InstallX86
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddInClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: InstallX86
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#AddInClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetCodeBase|x86}"; Check: InstallX86
Root: HKCU32; Subkey: "Software\Classes\{#AddInProgId}"; ValueType: string; ValueData: "Schnellbausteine für OneNote"; Flags: uninsdeletekey; Check: InstallX86
Root: HKCU32; Subkey: "Software\Classes\{#AddInProgId}\CLSID"; ValueType: string; ValueData: "{#AddInClsid}"; Check: InstallX86
Root: HKCU32; Subkey: "Software\Classes\AppID\{#AddInClsid}"; ValueType: string; ValueData: "Schnellbausteine für OneNote"; Flags: uninsdeletekey; Check: InstallX86
Root: HKCU32; Subkey: "Software\Classes\AppID\{#AddInClsid}"; ValueType: string; ValueName: "DllSurrogate"; ValueData: ""; Check: InstallX86
Root: HKCU32; Subkey: "Software\Microsoft\Office\OneNote\AddIns\{#AddInProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Schnellbausteine"; Flags: uninsdeletekey; Check: InstallX86
Root: HKCU32; Subkey: "Software\Microsoft\Office\OneNote\AddIns\{#AddInProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Lokale Schnellbausteine für OneNote Desktop"; Check: InstallX86
Root: HKCU32; Subkey: "Software\Microsoft\Office\OneNote\AddIns\{#AddInProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: InstallX86

; 64-bit OneNote registration.
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddInClsid}"; ValueType: string; ValueData: "Schnellbausteine für OneNote"; Flags: uninsdeletekey; Check: InstallX64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddInClsid}"; ValueType: string; ValueName: "AppID"; ValueData: "{#AddInClsid}"; Check: InstallX64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddInClsid}\ProgID"; ValueType: string; ValueData: "{#AddInProgId}"; Check: InstallX64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddInClsid}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"; Check: InstallX64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddInClsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Check: InstallX64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddInClsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "Schnellbausteine.AddIn.AddIn"; Check: InstallX64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddInClsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AddInAssembly}"; Check: InstallX64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddInClsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Check: InstallX64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#AddInClsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetCodeBase|x64}"; Check: InstallX64
Root: HKCU64; Subkey: "Software\Classes\{#AddInProgId}"; ValueType: string; ValueData: "Schnellbausteine für OneNote"; Flags: uninsdeletekey; Check: InstallX64
Root: HKCU64; Subkey: "Software\Classes\{#AddInProgId}\CLSID"; ValueType: string; ValueData: "{#AddInClsid}"; Check: InstallX64
Root: HKCU64; Subkey: "Software\Classes\AppID\{#AddInClsid}"; ValueType: string; ValueData: "Schnellbausteine für OneNote"; Flags: uninsdeletekey; Check: InstallX64
Root: HKCU64; Subkey: "Software\Classes\AppID\{#AddInClsid}"; ValueType: string; ValueName: "DllSurrogate"; ValueData: ""; Check: InstallX64
Root: HKCU64; Subkey: "Software\Microsoft\Office\OneNote\AddIns\{#AddInProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Schnellbausteine"; Flags: uninsdeletekey; Check: InstallX64
Root: HKCU64; Subkey: "Software\Microsoft\Office\OneNote\AddIns\{#AddInProgId}"; ValueType: string; ValueName: "Description"; ValueData: "Lokale Schnellbausteine für OneNote Desktop"; Check: InstallX64
Root: HKCU64; Subkey: "Software\Microsoft\Office\OneNote\AddIns\{#AddInProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Check: InstallX64

[Code]
var
  OneNotePlatform: String;

function NormalizePlatform(Value: String): String;
begin
  Value := Lowercase(Trim(Value));
  if (Value = 'x86') or (Value = '32') then
    Result := 'x86'
  else if (Value = 'x64') or (Value = '64') then
    Result := 'x64'
  else
    Result := '';
end;

function ReadPlatformFromRegistry(): String;
var
  Value: String;
  InstallPath: String;
begin
  Result := '';
  if IsWin64 and
     RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
       'Platform', Value) then
    Result := NormalizePlatform(Value);

  if (Result = '') and
     RegQueryStringValue(HKLM32, 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
       'Platform', Value) then
    Result := NormalizePlatform(Value);

  if Result <> '' then
    Exit;

  if IsWin64 and
     RegKeyExists(HKLM64, 'SOFTWARE\Microsoft\Office\16.0\OneNote\InstallRoot') then
  begin
    Result := 'x64';
    Exit;
  end;

  if RegKeyExists(HKLM32, 'SOFTWARE\Microsoft\Office\16.0\OneNote\InstallRoot') then
  begin
    Result := 'x86';
    Exit;
  end;

  if IsWin64 and
     RegQueryStringValue(HKLM64,
       'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\ONENOTE.EXE',
       '', InstallPath) then
  begin
    if Pos('\PROGRAM FILES (X86)\', Uppercase(InstallPath)) > 0 then
      Result := 'x86'
    else
      Result := 'x64';
    Exit;
  end;

  if RegQueryStringValue(HKLM32,
     'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\ONENOTE.EXE',
     '', InstallPath) then
    Result := 'x86';
end;

function IsOneNoteRunning(): Boolean;
var
  Locator: Variant;
  Services: Variant;
  Processes: Variant;
begin
  try
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    Services := Locator.ConnectServer('', 'root\CIMV2');
    Processes := Services.ExecQuery(
      'SELECT ProcessId FROM Win32_Process WHERE Name = "ONENOTE.EXE"');
    Result := Processes.Count > 0;
  except
    RaiseException(
      'Der Status von OneNote konnte nicht geprüft werden. ' +
      'Bitte schließen Sie OneNote und starten Sie Setup erneut.');
  end;
end;

function IsDotNet48Installed(): Boolean;
var
  Release: Cardinal;
begin
  Result :=
    RegQueryDWordValue(HKLM32,
      'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) and
    (Release >= 528040);
end;

function InitializeSetup(): Boolean;
var
  OverridePlatform: String;
begin
  Result := False;
  if not IsDotNet48Installed() then
  begin
    MsgBox(
      'Microsoft .NET Framework 4.8 ist erforderlich. ' +
      'Bitte installieren Sie es über Windows Update und starten Sie Setup erneut.',
      mbCriticalError, MB_OK);
    Exit;
  end;

  OverridePlatform := NormalizePlatform(
    ExpandConstant('{param:ONENOTEBITNESS|}'));
  if OverridePlatform <> '' then
    OneNotePlatform := OverridePlatform
  else
    OneNotePlatform := ReadPlatformFromRegistry();

  if OneNotePlatform = '' then
  begin
    MsgBox(
      'Es wurde keine unterstützte OneNote-Desktop-Installation gefunden.'#13#10#13#10 +
      'Unterstützt werden OneNote aus Microsoft 365 und OneNote 2016. ' +
      'OneNote für Windows 10, Web und Store werden nicht unterstützt.',
      mbCriticalError, MB_OK);
    Exit;
  end;

  if (OneNotePlatform = 'x64') and not IsWin64 then
  begin
    MsgBox('64-Bit-OneNote kann auf diesem Windows-System nicht installiert sein.',
      mbCriticalError, MB_OK);
    Exit;
  end;

  Result := True;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  if IsOneNoteRunning() then
    Result :=
      'OneNote ist noch geöffnet. Schließen Sie alle OneNote-Fenster vollständig ' +
      'und klicken Sie anschließend erneut auf Installieren.';
end;

function InitializeUninstall(): Boolean;
begin
  Result := not IsOneNoteRunning();
  if not Result then
    MsgBox(
      'OneNote ist noch geöffnet. Schließen Sie OneNote vollständig und ' +
      'starten Sie die Deinstallation danach erneut.',
      mbCriticalError, MB_OK);
end;

function InstallX86(): Boolean;
begin
  Result := OneNotePlatform = 'x86';
end;

function InstallX64(): Boolean;
begin
  Result := OneNotePlatform = 'x64';
end;

function GetCodeBase(Param: String): String;
var
  FileName: String;
begin
  FileName := ExpandConstant('{app}\' + Param +
    '\Schnellbausteine.AddIn.dll');
  StringChangeEx(FileName, '\', '/', True);
  StringChangeEx(FileName, ' ', '%20', True);
  Result := 'file:///' + FileName;
end;

function UpdateReadyMemo(
  Space, NewLine, MemoUserInfoInfo, MemoDirInfo, MemoTypeInfo,
  MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
var
  BitnessText: String;
begin
  if OneNotePlatform = 'x64' then
    BitnessText := '64 Bit'
  else
    BitnessText := '32 Bit';

  Result :=
    'Setup ist bereit, Schnellbausteine für OneNote zu installieren.' + NewLine +
    NewLine +
    'Erkannte OneNote-Version: ' + BitnessText + NewLine +
    'Installation: Nur für den aktuellen Benutzer (keine Administratorrechte)' +
    NewLine + NewLine +
    'Klicken Sie auf Installieren, um fortzufahren.';
end;
