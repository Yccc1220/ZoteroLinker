# Zotero Linker Installer

Builds the released Microsoft Word VSTO add-in installer with Inno Setup 6.

## Prerequisites

- Inno Setup 6+
- Visual Studio 2022 with Office/SharePoint development workload
- .NET Framework 4.7.2 targeting pack
- Visual Studio Tools for Office Runtime on target machines

## Build

```batch
cd installer
build.bat 1.0.0 Release
```

Output:

```text
dist\ZoteroLinkerSetup-1.0.0.exe
```

## What the installer does

1. Checks that the VSTO Runtime 10.0 is installed.
2. Copies the Word VSTO files to the selected installation directory.
3. Installs the VSTO signing certificate into Root and Trusted Publisher.
4. Registers Zotero Linker under the Word add-in registry keys.
5. Uses `file:///.../Zotero-linker.vsto|vstolocal` for the VSTO manifest.
6. Enables machine-level VSTO loading with `EnableLocalMachineVSTO=1`.
7. Writes VSTO security inclusion entries for the installing user and machine.
8. Cleans old Zotero Linker VSTO registry/cache data during upgrade and uninstall.
