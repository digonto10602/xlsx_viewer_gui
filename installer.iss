[Setup]
AppName=Xlsx Row Viewer
AppVersion=1.0.0
DefaultDirName={pf}\XlsxRowViewer
DefaultGroupName=Xlsx Row Viewer
OutputBaseFilename=XlsxRowViewerSetup
Compression=lzma
SolidCompression=yes
DisableProgramGroupPage=yes

[Tasks]
Name: "desktopicon"; Description: "Create a Desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "dist\XlsxRowViewer.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "sample.xlsx"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Xlsx Row Viewer"; Filename: "{app}\XlsxRowViewer.exe"
Name: "{group}\README"; Filename: "{app}\README.txt"
Name: "{group}\Sample Excel File"; Filename: "{app}\sample.xlsx"
Name: "{commondesktop}\Xlsx Row Viewer"; Filename: "{app}\XlsxRowViewer.exe"; Tasks: desktopicon

; Right-click integration: Right click .xlsx -> "Open with Xlsx Row Viewer"
[Registry]
Root: HKCR; Subkey: "SystemFileAssociations\.xlsx\shell\Open with Xlsx Row Viewer"; ValueType: string; ValueName: ""; ValueData: "Open with Xlsx Row Viewer"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\.xlsx\shell\Open with Xlsx Row Viewer\command"; ValueType: string; ValueName: ""; ValueData: """{app}\XlsxRowViewer.exe"" ""%1"""; Flags: uninsdeletekey

[Run]
Filename: "{app}\XlsxRowViewer.exe"; Description: "Run Xlsx Row Viewer now"; Flags: nowait postinstall skipifsilent