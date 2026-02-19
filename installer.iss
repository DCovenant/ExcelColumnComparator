[Setup]
AppName=Excel Column Comparator
AppVersion=1.0.0
AppPublisher=Excel Tools
AppPublisherURL=https://github.com
DefaultDirName={autopf}\ExcelComparator
DefaultGroupName=Excel Comparator
OutputDir=dist\installer
OutputBaseFilename=ExcelColumnComparator-Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
UsePreviousAppDir=yes
WizardStyle=modern
UninstallDisplayIcon={app}\union_gui.exe

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; Flags: unchecked

[Files]
Source: "dist\union_gui.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Excel Column Comparator"; Filename: "{app}\union_gui.exe"
Name: "{group}\{cm:UninstallProgram,Excel Column Comparator}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\Excel Column Comparator"; Filename: "{app}\union_gui.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\union_gui.exe"; Description: "{cm:LaunchProgram,Excel Column Comparator}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: dirifempty; Name: "{app}"
