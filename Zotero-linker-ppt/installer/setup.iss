; Zotero Linker PPT - Inno Setup Installer
; Build: iscc /DVersion=1.0.0 /DConfig=Release setup.iss

#define AppName "Zotero Linker PPT"
#define AppPublisher "Yccc1220"
#define AppUrl "https://github.com/Yccc1220/ZoteroLinker"
#define AddInName "Zotero-linker-ppt"
#define AddInDescription "Zotero citation integration for PowerPoint"

#ifndef Version
  #define Version "1.0.0"
#endif
#ifndef Config
  #define Config "Release"
#endif

[Setup]
AppId={{E03AF355-404C-4E8E-A9E1-9AD4B2FA214D}
AppName={#AppName}
AppVersion={#Version}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppUrl}
DefaultDirName={commonpf}\Zotero Linker PPT
DefaultGroupName=Zotero Linker PPT
DisableProgramGroupPage=yes
OutputDir=..\dist
OutputBaseFilename=zoterolinkerppt-{#Version}
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\PowerPoint\icon.ico

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "chinesesimplified"; MessagesFile: "ChineseSimplified.isl"

[Files]
Source: "..\bin\{#Config}\Zotero-linker-ppt.vsto"; DestDir: "{app}\PowerPoint"; Flags: ignoreversion
Source: "..\bin\{#Config}\Zotero-linker-ppt.dll"; DestDir: "{app}\PowerPoint"; Flags: ignoreversion
Source: "..\bin\{#Config}\Zotero-linker-ppt.dll.manifest"; DestDir: "{app}\PowerPoint"; Flags: ignoreversion
Source: "..\bin\{#Config}\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"; DestDir: "{app}\PowerPoint"; Flags: ignoreversion
Source: "vsto-signing.cer"; DestDir: "{app}"; Flags: ignoreversion
Source: "WriteVstoInclusions.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "WritePowerPointAddinRegistration.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "WriteWpsPresentationWhitelist.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\tools\ForceClean.ps1"; DestDir: "{tmp}"; Flags: dontcopy
Source: "..\tools\ForceClean.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "icon.ico"; DestDir: "{app}\PowerPoint"; Flags: ignoreversion

[Registry]
; PowerPoint Add-in: versionless path
Root: HKLM; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|PowerPoint\{#AddInName}.vsto}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

; PowerPoint Add-in: Office 16.0 path
Root: HKLM; Subkey: "Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|PowerPoint\{#AddInName}.vsto}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey

; 32-bit Office on 64-bit Windows
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|PowerPoint\{#AddInName}.vsto}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|PowerPoint\{#AddInName}.vsto}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey; Check: IsWin64

; ClickToRun virtualized registry for Microsoft 365 / Office C2R
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|PowerPoint\{#AddInName}.vsto}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|PowerPoint\{#AddInName}.vsto}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|PowerPoint\{#AddInName}.vsto}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Description"; ValueData: "{#AddInDescription}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:GetManifestUri|PowerPoint\{#AddInName}.vsto}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKLM; Subkey: "Software\WOW6432Node\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\16.0\PowerPoint\Addins\{#AddInName}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "1"; Flags: uninsdeletekey; Check: IsWin64

; Let VSTO load machine-level Office add-ins.
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; ValueType: string; ValueName: "EnableLocalMachineVSTO"; ValueData: "1"; Flags: uninsdeletevalue

; WPS Presentation add-in whitelist. WPS stores allowed add-ins as values under AddinsWL.
Root: HKLM; Subkey: "Software\Kingsoft\Office\WPP\AddinsWL"; ValueType: string; ValueName: "{#AddInName}"; ValueData: ""; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\WOW6432Node\Kingsoft\Office\WPP\AddinsWL"; ValueType: string; ValueName: "{#AddInName}"; ValueData: ""; Flags: uninsdeletevalue; Check: IsWin64
Root: HKLM; Subkey: "Software\Kingsoft\Office\WPP\AddinsWL"; ValueType: string; ValueName: "{#AppPublisher}.{#AppName}"; ValueData: ""; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\WOW6432Node\Kingsoft\Office\WPP\AddinsWL"; ValueType: string; ValueName: "{#AppPublisher}.{#AppName}"; ValueData: ""; Flags: uninsdeletevalue; Check: IsWin64

[Run]
Filename: "{sys}\certutil.exe"; Parameters: "-addstore -f ""Root"" ""{app}\vsto-signing.cer"""; Flags: runhidden
Filename: "{sys}\certutil.exe"; Parameters: "-addstore -f ""TrustedPublisher"" ""{app}\vsto-signing.cer"""; StatusMsg: "{cm:InstallingCertificate}"; Flags: runhidden
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\WriteVstoInclusions.ps1"" -ManifestPath ""{app}\PowerPoint\{#AddInName}.vsto"" -Target HKLM"; StatusMsg: "{cm:RegisteringPowerPoint}"; Flags: runhidden
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\WriteVstoInclusions.ps1"" -ManifestPath ""{app}\PowerPoint\{#AddInName}.vsto"" -Target HKCU"; StatusMsg: "{cm:RegisteringPowerPoint}"; Flags: runhidden runasoriginaluser
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\WritePowerPointAddinRegistration.ps1"" -AddInName ""{#AddInName}"" -ManifestPath ""{app}\PowerPoint\{#AddInName}.vsto"" -FriendlyName ""{#AppName}"" -Description ""{#AddInDescription}"""; StatusMsg: "{cm:RegisteringPowerPoint}"; Flags: runhidden runasoriginaluser
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\WriteWpsPresentationWhitelist.ps1"" -Name ""{#AddInName}"""; StatusMsg: "{cm:RegisteringWps}"; Flags: runhidden runasoriginaluser
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\WriteWpsPresentationWhitelist.ps1"" -Name ""{#AppPublisher}.{#AppName}"""; StatusMsg: "{cm:RegisteringWps}"; Flags: runhidden runasoriginaluser

[CustomMessages]
english.InstallingCertificate=Installing add-in certificate...
english.RegisteringPowerPoint=Registering Zotero Linker PPT for PowerPoint...
english.RegisteringWps=Registering Zotero Linker PPT for WPS Presentation...
chinesesimplified.InstallingCertificate=Installing add-in certificate...
chinesesimplified.RegisteringPowerPoint=Registering Zotero Linker PPT for PowerPoint...
chinesesimplified.RegisteringWps=Registering Zotero Linker PPT for WPS Presentation...

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
