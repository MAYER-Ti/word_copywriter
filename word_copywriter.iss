[Setup]
AppName=Word Copywriter
AppVersion=1.0
DefaultDirName={autopf}\Word Copywriter
DefaultGroupName=Word Copywriter
OutputDir=installer
OutputBaseFilename=WordCopywriterSetup
Compression=lzma2
SolidCompression=yes
SetupIconFile=resources\icon.ico       ; иконка самого установщика

[Files]
Source: "dist\word_copywriter.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "resources\icon.ico";         DestDir: "{app}\resources"; Flags: ignoreversion
Source: "dist\templates\*";           DestDir: "{app}\templates"; Flags: recursesubdirs ignoreversion
Source: "dist\resources\*";           DestDir: "{app}\resources"; Flags: recursesubdirs ignoreversion
; ...

[Icons]
Name: "{group}\Word Copywriter"; Filename: "{app}\word_copywriter.exe"; IconFilename: "{app}\resources\icon.ico"
Name: "{userdesktop}\Word Copywriter"; Filename: "{app}\word_copywriter.exe"; IconFilename: "{app}\resources\icon.ico"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"
