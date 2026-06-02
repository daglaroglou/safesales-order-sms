#define MyAppVersion "1.0.0"
#define RuntimeInstallerSource "..\packaging\prereqs\WindowsAppRuntimeInstall-x64.exe"
#ifexist RuntimeInstallerSource
  #define HasRuntimeInstaller
#endif

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
SetupIconFile=..\assets\SafeSalesSMS.ico
UninstallDisplayIcon={app}\SafeSalesSMS.exe
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "..\dist\SafeSalesSMS.exe"; DestDir: "{app}"; Flags: ignoreversion
#ifdef HasRuntimeInstaller
Source: "{#RuntimeInstallerSource}"; DestDir: "{tmp}"; Flags: deleteafterinstall
#endif

[Icons]
Name: "{group}\SafeSales Order SMS"; Filename: "{app}\SafeSalesSMS.exe"
Name: "{userdesktop}\SafeSales Order SMS"; Filename: "{app}\SafeSalesSMS.exe"; Tasks: desktopicon

[Run]
#ifdef HasRuntimeInstaller
Filename: "{tmp}\WindowsAppRuntimeInstall-x64.exe"; Parameters: "/quiet /norestart"; StatusMsg: "Installing Windows App Runtime..."; Flags: waituntilterminated runhidden
#endif
Filename: "{app}\SafeSalesSMS.exe"; Description: "Launch SafeSales Order SMS"; Flags: nowait postinstall skipifsilent
