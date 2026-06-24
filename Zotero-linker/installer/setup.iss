; Zotero Linker - Inno Setup Installer
; Build: iscc /DVersion=1.0.0 /DConfig=Release setup.iss

#define AppName "Zotero Linker"
#define AppPublisher "Yccc1220"
#define AppUrl "https://github.com/Yccc1220/ZoteroLinker"
#define AddInName "Zotero-linker"
#define AddInDescription "Zotero citation linker for Word"

#ifndef Version
  #define Version "1.0.0"
#endif
#ifndef Config
  #define Config "Release"
#endif

[Setup]
AppId={{FAEE369B-CA0E-4BCD-8635-9F443F3E16FB}
AppName={#AppName}
AppVersion={#Version}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppUrl}
DefaultDirName={commonpf}\Zotero Linker
DefaultGroupName=Zotero Linker
DisableProgramGroupPage=yes
OutputDir=..\dist
OutputBaseFilename=ZoteroLinkerSetup-{#Version}
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\Word\icon.ico

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "chinesesimplified"; MessagesFile: "ChineseSimplified.isl"

[Files]
Source: "..\bin\{#Config}\Zotero-linker.vsto"; DestDir: "{app}\Word"; Flags: ignoreversion
Source: "..\bin\{#Config}\Zotero-linker.dll"; DestDir: "{app}\Word"; Flags: ignoreversion
Source: "..\bin\{#Config}\Zotero-linker.dll.manifest"; DestDir: "{app}\Word"; Flags: ignoreversion
Source: "..\bin\{#Config}\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"; DestDir: "{app}\Word"; Flags: ignoreversion
Source: "vsto-signing.cer"; DestDir: "{app}"; Flags: ignoreversion
Source: "WriteVstoInclusions.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "WriteWpsWhitelist.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\tools\ForceClean.ps1"; DestDir: "{tmp}"; Flags: dontcopy
Source: "..\tools\ForceClean.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon.ico"; DestDir: "{app}\Word"; Flags: ignoreversion

[Registry]
; Word Add-in: versionless path
Root: HKLM; Subkey: "Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|Word\{#AddInName}.vsto}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

; Word Add-in: Office 16.0 path
Root: HKLM; Subkey: "Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|Word\{#AddInName}.vsto}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

; 32-bit Office on 64-bit Windows
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|Word\{#AddInName}.vsto}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|Word\{#AddInName}.vsto}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey; Check: IsWin64

; ClickToRun virtualized registry for Microsoft 365 / Office C2R
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|Word\{#AddInName}.vsto}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|Word\{#AddInName}.vsto}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|Word\{#AddInName}.vsto}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|Word\{#AddInName}.vsto}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\Word\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey; Check: IsWin64

; Let VSTO load machine-level Office add-ins.
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; ValueType: string; ValueName: "EnableLocalMachineVSTO"; ValueData: "1"; Flags: uninsdeletevalue

; WPS Writer add-in whitelist. Equivalent to Advanced Installer's [Manufacturer].[ProductName].
Root: HKLM; Subkey: "Software\Kingsoft\Office\wps\AddinsWL\{#AppPublisher}.{#AppName}"; Flags: uninsdeletekey

[Run]
Filename: "{sys}\certutil.exe"; Parameters: "-addstore -f ""Root"" ""{app}\vsto-signing.cer"""; Flags: runhidden
Filename: "{sys}\certutil.exe"; Parameters: "-addstore -f ""TrustedPublisher"" ""{app}\vsto-signing.cer"""; StatusMsg: "{cm:InstallingCertificate}"; Flags: runhidden
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\WriteVstoInclusions.ps1"" -ManifestPath ""{app}\Word\{#AddInName}.vsto"" -Target HKLM"; StatusMsg: "{cm:RegisteringWord}"; Flags: runhidden
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\WriteVstoInclusions.ps1"" -ManifestPath ""{app}\Word\{#AddInName}.vsto"" -Target HKCU"; StatusMsg: "{cm:RegisteringWord}"; Flags: runhidden runasoriginaluser
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\WriteWpsWhitelist.ps1"" -Name ""{#AppPublisher}.{#AppName}"""; StatusMsg: "{cm:RegisteringWps}"; Flags: runhidden runasoriginaluser

[CustomMessages]
english.InstallingCertificate=Installing add-in certificate...
english.RegisteringWord=Registering Zotero Linker for Word...
english.RegisteringWps=Registering Zotero Linker for WPS Writer...
chinesesimplified.InstallingCertificate=Installing add-in certificate...
chinesesimplified.RegisteringWord=Registering Zotero Linker for Word...
chinesesimplified.RegisteringWps=Registering Zotero Linker for WPS Writer...

[Code]
function VstoInstallerExists: Boolean;
var
  Path: string;
begin
  Path := ExpandConstant('{commonpf32}\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe');
  if FileExists(Path) then
  begin
    Result := True;
    Exit;
  end;
  Path := ExpandConstant('{commonpf}\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe');
  Result := FileExists(Path);
end;

function InitializeSetup: Boolean;
begin
  if not VstoInstallerExists then
  begin
    SuppressibleMsgBox(
      'Microsoft Visual Studio Tools for Office Runtime is required but was not found.'#13#13 +
      'Please install the VSTO Runtime before installing this add-in.'#13#13 +
      'Download: https://go.microsoft.com/fwlink/?LinkId=140384',
      mbCriticalError, MB_OK, 0);
    Result := False;
  end
  else
    Result := True;
end;

function GetManifestUri(Param: string): string;
var
  AppDir: string;
begin
  AppDir := ExpandConstant('{app}');
  StringChange(AppDir, '\', '/');
  StringChange(Param, '\', '/');
  Result := 'file:///' + AppDir + '/' + Param + '|vstolocal';
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ResultCode: Integer;
begin
  if CurUninstallStep = usUninstall then
  begin
    Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
         '-ExecutionPolicy Bypass -File "' + ExpandConstant('{app}') + '\ForceClean.ps1" -InstallRoot "' + ExpandConstant('{app}') + '"',
         '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    Log('ForceClean exited with code ' + IntToStr(ResultCode));
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ResultCode: Integer;
begin
  Result := '';
  ExtractTemporaryFile('ForceClean.ps1');
  Exec(ExpandConstant('{sys}\WindowsPowerShell\v1.0\powershell.exe'),
       '-ExecutionPolicy Bypass -File "' + ExpandConstant('{tmp}') + '\ForceClean.ps1" -InstallRoot "' + ExpandConstant('{app}') + '" -RemoveInstallDir',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Log('ForceClean exited with code ' + IntToStr(ResultCode));
end;
