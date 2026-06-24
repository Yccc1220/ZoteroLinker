# Release Assets

This folder follows the LaTeXSnipper Office plugin release style: the Windows VSTO installer is built locally, committed here with a SHA-256 checksum, and GitHub Actions validates and uploads it to GitHub Releases.

To create a new release:

1. Build the installer locally:

   ```batch
   cd Zotero-linker\installer
   build.bat 1.0.0 Release
   ```

2. Copy the installer to this folder:

   ```text
   release\ZoteroLinkerSetup-1.0.0.exe
   ```

3. Generate the checksum file:

   ```powershell
   $hash = (Get-FileHash .\release\ZoteroLinkerSetup-1.0.0.exe -Algorithm SHA256).Hash.ToLowerInvariant()
   Set-Content .\release\ZoteroLinkerSetup-1.0.0.exe.sha256 "$hash  ZoteroLinkerSetup-1.0.0.exe"
   ```

4. Commit the installer and checksum, then push a tag such as `v1.0.0`.
