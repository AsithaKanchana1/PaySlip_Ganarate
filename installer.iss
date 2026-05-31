; ============================================================
;  installer.iss  —  Inno Setup Script
;  Creates a professional Windows installer (.exe) for
;  New Lanka Clothing Pay Slip Generator
;
;  How to build the installer:
;    1. Install Inno Setup: https://jrsoftware.org/isdl.php
;    2. First run build_windows.bat to create the dist\ folder
;    3. Right-click this file → Compile  (or open in Inno Setup IDE)
;    4. The installer will appear as:  Output\PaySlipGenerator_Setup.exe
; ============================================================

#define AppName      "Pay Slip Generator"
#define AppVersion   "1.0"
#define AppPublisher "New Lanka Clothing"
#define AppURL       "https://github.com/AsithaKanchana1/PaySlip_Ganarate"
#define AppExeName   "PaySlipGenerator.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
DefaultDirName={autopf}\NewLankaClothing\PaySlipGenerator
DefaultGroupName=New Lanka Clothing
AllowNoIcons=yes
LicenseFile=
OutputDir=Output
OutputBaseFilename=PaySlipGenerator_Setup_v{#AppVersion}
#if FileExists("icon.ico")
SetupIconFile=icon.ico
WizardSmallImageFile=icon.ico
#endif
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#AppExeName}
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ChangesAssociations=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";    Description: "Create a &Desktop shortcut";          GroupDescription: "Additional icons:"; Flags: unchecked
Name: "startmenuicon";  Description: "Create a &Start Menu shortcut";        GroupDescription: "Additional icons:"

[Files]
; Main application files from PyInstaller output
Source: "dist\PaySlipGenerator\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; Create the Excel folder inside the install directory
Source: "Excel\.gitkeep"; DestDir: "{app}\Excel"; Flags: ignoreversion

; README
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
; Ensure Excel folder exists and is writable
Name: "{app}\Excel"

[Icons]
; Start Menu
Name: "{group}\{#AppName}";           Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\{#AppExeName}"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"

; Desktop (optional)
Name: "{autodesktop}\{#AppName}";     Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
; Offer to launch after install
Filename: "{app}\{#AppExeName}"; Description: "Launch Pay Slip Generator now"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Clean up any PDFs generated in the app folder
Type: filesandordirs; Name: "{app}\*.pdf"

[Messages]
WelcomeLabel2=This will install [name/ver] on your computer.%n%nNew Lanka Clothing Pay Slip Generator lets you generate printer-ready PDF pay slips for all employees from an Excel file.%n%nClick Next to continue.

[CustomMessages]
english.FinishedLabel=Setup has finished installing [name] on your computer.%n%nTo get started:%n  1. Launch Pay Slip Generator%n  2. Copy your Excel salary file into the Excel\ folder inside the installation directory%n  3. Select the file in the app and click Generate

[Code]
// Show a helpful message after install pointing to the Excel folder
procedure CurStepChanged(CurStep: TSetupStep);
var
  ExcelDir: String;
begin
  if CurStep = ssDone then
  begin
    ExcelDir := ExpandConstant('{app}\Excel');
    MsgBox(
      'Installation complete!' + #13#10 + #13#10 +
      'To use the Pay Slip Generator:' + #13#10 +
      '  1. Copy your Excel salary file (Slary_Slips.xlsx) to:' + #13#10 +
      '     ' + ExcelDir + #13#10 + #13#10 +
      '  2. Launch the app from your Start Menu or Desktop.' + #13#10 +
      '  3. Browse for the Excel file, choose the month, and click Generate!',
      mbInformation, MB_OK
    );
  end;
end;
