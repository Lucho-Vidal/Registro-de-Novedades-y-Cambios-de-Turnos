; Script de Inno Setup optimizado para RENO (Estándar CRONO-PCT).
; Compresión LZMA2 Ultra64 con diccionario de 64 MB para reducción extrema (<42 MB)
; Compatible con instalación desatendida y auto-updater silencioso sin permisos administrativos.

#define MyAppName "RENO"
#define MyAppVersion "2.0.0"
#define MyAppPublisher "Trenes Argentinos - luvidev"
#define MyAppURL "https://luciano-vidal.web.app/"
#define MyAppExeName "RENO.exe"
#define MyAppAssocName MyAppName + " File"
#define MyAppAssocExt ".myp"
#define MyAppAssocKey StringChange(MyAppAssocName, " ", "") + MyAppAssocExt

[Setup]
AppId={{F3DAD1DB-89B6-4C1E-BB17-11A2CE031336}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}

ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
ChangesAssociations=yes
DisableProgramGroupPage=yes

LicenseFile=LICENSE.txt
InfoBeforeFile=beforeInstalation.txt
InfoAfterFile=afterInstalation.txt

; Instalación sin elevación obligatoria de privilegios (Standard User)
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=Instalador
OutputBaseFilename=RENO_Setup_v{#MyAppVersion}

; --- COMPRESIÓN ULTRA64 (DICCIONARIO 64 MB) ESTÁNDAR CRONO ---
; Reduce el binario final a ~38-42 MB para respetar el límite de 50 MB de Supabase Storage
Compression=lzma2/ultra64
InternalCompressLevel=ultra64
SolidCompression=yes
LZMAUseSeparateProcess=yes
LZMANumBlockThreads=4

WizardStyle=modern

[Code]
const
  FECHA_VENCIMIENTO = '20271231';

function FormatearFecha(const Fecha: String): String;
begin
  Result := Copy(Fecha, 7, 2) + '/' +
            Copy(Fecha, 5, 2) + '/' +
            Copy(Fecha, 1, 4);
end;

function InitializeSetup(): Boolean;
begin
  Result := CompareStr(
    GetDateTimeString('yyyymmdd', #0, #0),
    FECHA_VENCIMIENTO
  ) <= 0;

  if not Result then
  begin
    MsgBox(
      'Este instalador ha vencido.' + #13#10 +
      'Fecha límite: ' + FormatearFecha(FECHA_VENCIMIENTO),
      mbError,
      MB_OK
    );
  end;
end;

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Archivos auxiliares opcionales
Source: "LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "beforeInstalation.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "afterInstalation.txt"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocExt}\OpenWithProgids"; ValueType: string; ValueName: "{#MyAppAssocKey}"; ValueData: ""; Flags: uninsdeletevalue
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}"; ValueType: string; ValueName: ""; ValueData: "{#MyAppAssocName}"; Flags: uninsdeletekey
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#MyAppExeName},0"
Root: HKA; Subkey: "Software\Classes\{#MyAppAssocKey}\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExeName}\SupportedTypes"; ValueType: string; ValueName: ".myp"; ValueData: ""

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Sin skipifsilent para que la actualización automática silenciosa (/SILENT) relance RENO al finalizar
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall
