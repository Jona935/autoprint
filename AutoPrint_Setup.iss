#define AppName      "AutoPrint"
#define AppVersion   "1.2"
#define AppPublisher "AutoPrint"
#define AppExeName   "AutoPrint.exe"
#define AppDesc      "Impresion automatica de PDFs desde Google Drive"

[Setup]
AppId={{B3F2A1C4-9D8E-4F5A-A2B7-C6D3E0F1A8B9}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} v{#AppVersion}
AppPublisherURL=https://github.com/Jona935/autoprint
AppSupportURL=https://github.com/Jona935/autoprint
AppUpdatesURL=https://github.com/Jona935/autoprint/releases
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
LicenseFile=
OutputDir=.\installer
OutputBaseFilename=AutoPrint_Setup_v{#AppVersion}
SetupIconFile=autoprint.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName}
VersionInfoVersion={#AppVersion}.0.0
VersionInfoDescription={#AppDesc}
ShowLanguageDialog=no

; Cierre automatico de la app si esta corriendo al instalar/actualizar
CloseApplications=force
RestartApplications=no

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon";  Description: "Crear icono en el Escritorio";       GroupDescription: "Iconos adicionales:"; Flags: unchecked
Name: "startupicon";  Description: "Iniciar automaticamente con Windows"; GroupDescription: "Opciones de inicio:";  Flags: unchecked

[Files]
; ignoreversion + replacesameversion asegura que siempre se reemplaza el exe aunque tenga la misma version
Source: "dist\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "autoprint.ico";      DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}";             Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\autoprint.ico"
Name: "{group}\Desinstalar {#AppName}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#AppName}";     Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\autoprint.ico"; Tasks: desktopicon

[Registry]
; Inicio automatico con Windows (solo si el usuario lo eligio en las Tasks)
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"; \
  ValueType: string; ValueName: "{#AppName}"; \
  ValueData: """{app}\{#AppExeName}"""; \
  Flags: uninsdeletevalue; Tasks: startupicon

; Si NO se eligio startupicon, borrar la entrada de autostart de versiones anteriores
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"; \
  ValueName: "{#AppName}"; \
  Flags: deletevalue uninsdeletevalue; Tasks: not startupicon

[Run]
Filename: "{app}\{#AppExeName}"; \
  Description: "Iniciar {#AppName} ahora"; \
  Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "taskkill"; Parameters: "/F /IM {#AppExeName}"; Flags: runhidden; RunOnceId: "KillApp"

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Code]
procedure KillRunningApp();
var
  ResultCode: Integer;
begin
  // Matar via taskkill y esperar a que termine
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM {#AppExeName}', '',
       SW_HIDE, ewWaitUntilTerminated, ResultCode);
  // Pausa breve para que Windows libere el archivo
  Sleep(1500);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
    KillRunningApp();
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
    KillRunningApp();
end;
