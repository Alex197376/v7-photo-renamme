; Script Inno Setup pour Photo Manager Pro
[Setup]
AppName=Photo Manager Pro
AppVersion=6.2
DefaultDirName={pf}\Photo Manager Pro
; Script Inno Setup pour Photo Manager Pro
[Setup]
AppName=Photo Manager Pro
AppVersion=6.2
DefaultDirName={pf}\Photo Manager Pro
DefaultGroupName=Photo Manager Pro
OutputDir=.
OutputBaseFilename=Install_PhotoManagerPro
Compression=lzma
SolidCompression=yes
DisableWelcomePage=no
WizardStyle=modern
DefaultDialogFontName=Segoe UI


SetupIconFile="Icone V6_2.ico"

[Files]
[Files]
Source: "dist\Renommage_Photo_V6.1.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "button_presets.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "Icone V6_2.ico"; DestDir: "{app}"; Flags: ignoreversion
; Vous pouvez ajouter ici config.json ou d'autres fichiers si nécessaire

[Icons]
[Icons]
Name: "{group}\Photo Manager Pro"; Filename: "{app}\Renommage_Photo_V6.1.exe"; IconFilename: "{app}\Icone V6_2.ico"
Name: "{commondesktop}\Photo Manager Pro"; Filename: "{app}\Renommage_Photo_V6.1.exe"; IconFilename: "{app}\Icone V6_2.ico"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Créer un raccourci sur le bureau"; GroupDescription: "Raccourcis :"

[Run]
Filename: "{app}\Renommage_Photo_V6_1.exe"; Description: "Lancer Photo Manager Pro"; Flags: nowait postinstall skipifsilent
