[Setup]
AppName=CE Tool
AppVersion={#MyAppVersion}
DefaultDirName={pf}\CE Tool
OutputDir=Output
OutputBaseFilename=CE_Tool_Installer_{#MyAppVersion}
Compression=lzma
SolidCompression=yes
DisableProgramGroupPage=no
UninstallDisplayIcon={app}\ce_tool.exe

[Files]
; 包含 PyInstaller 產出的 dist/run 資料內容（遞迴）
; 包含 PyInstaller 產出的 dist/run 資料內容（遞迴）
; .iss 位於 installer/，而 dist/ce_tool 在 repo 根目錄下，因此使用 ../dist 路徑
Source: "..\dist\ce_tool\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Dirs]
; 建立一個空的資料目錄供 future database 存放（不會把使用者資料打包進去）
Name: "{commonappdata}\CE_Tool_Data"; Flags: uninsalwaysuninstall

[Icons]
Name: "{group}\CE Tool"; Filename: "{app}\ce_tool.exe"
Name: "{userdesktop}\CE Tool"; Filename: "{app}\ce_tool.exe"
