#define MyAppVersion "1.0.0"

[Setup]
AppName=SafeSales Order SMS
AppVersion={#MyAppVersion}
DefaultDirName={pf}\SafeSales Order SMS
DefaultGroupName=SafeSales Order SMS
OutputDir=release
OutputBaseFilename=SafeSalesSMSSetup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes
DisableProgramGroupPage=no
DisableReadyPage=no
DisableWelcomePage=no
InfoBeforeFile=WindowsAppRuntimeRequired.txt
LicenseFile=..\LICENSE
UninstallDisplayIcon={app}\SafeSalesSMS.exe
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "..\dist\SafeSalesSMS.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\SafeSales Order SMS"; Filename: "{app}\SafeSalesSMS.exe"
Name: "{userdesktop}\SafeSales Order SMS"; Filename: "{app}\SafeSalesSMS.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\SafeSalesSMS.exe"; Description: "Launch SafeSales Order SMS"; Flags: nowait postinstall skipifsilent
