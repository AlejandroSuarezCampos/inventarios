; Instalador de InventarioPC (Inno Setup 6)
; Compilar:  ISCC.exe /DMyAppVersion=1.1.0 installer\InventarioPC.iss
; Requiere que exista dist\InventarioPC.exe (pyinstaller InventarioPC.spec)

#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif
#define MyAppName "Inventario de Equipos"
#define MyAppExe "InventarioPC.exe"

[Setup]
AppId={{6F1E2B7C-4A3D-4C61-9B0E-8D2F5A7C3E91}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher=InventarioPC
VersionInfoVersion={#MyAppVersion}
DefaultDirName={autopf}\InventarioPC
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; Se instala sin permisos de administrador (para el usuario actual); se puede elegir «para todos»
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog commandline
OutputDir=..\dist
OutputBaseFilename=InventarioPC-Setup-{#MyAppVersion}
SetupIconFile=..\assets\icono.ico
UninstallDisplayIcon={app}\{#MyAppExe}
UninstallDisplayName={#MyAppName}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
ChangesAssociations=yes
CloseApplications=yes

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "escritorio"; Description: "Crear un acceso directo en el escritorio"; GroupDescription: "Accesos directos:"
Name: "abrircon"; Description: "Añadir «Abrir con Inventario de Equipos» a los archivos Excel (.xlsx)"; GroupDescription: "Integración con Windows:"

[Files]
Source: "..\dist\{#MyAppExe}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"
Name: "{group}\Servidor de inventario"; Filename: "{app}\{#MyAppExe}"; Parameters: "--servidor"; Comment: "Pone este equipo como servidor central del inventario"
Name: "{group}\Desinstalar {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExe}"; Tasks: escritorio

[Registry]
; «Abrir con» para .xlsx sin quitarle a Excel la asociación por defecto
Root: HKA; Subkey: "Software\Classes\InventarioPC.Importar"; ValueType: string; ValueName: ""; ValueData: "Importar en Inventario de Equipos"; Flags: uninsdeletekey; Tasks: abrircon
Root: HKA; Subkey: "Software\Classes\InventarioPC.Importar\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#MyAppExe},0"; Tasks: abrircon
Root: HKA; Subkey: "Software\Classes\InventarioPC.Importar\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExe}"" ""%1"""; Tasks: abrircon
Root: HKA; Subkey: "Software\Classes\.xlsx\OpenWithProgids"; ValueType: string; ValueName: "InventarioPC.Importar"; ValueData: ""; Flags: uninsdeletevalue; Tasks: abrircon
Root: HKA; Subkey: "Software\Classes\Applications\{#MyAppExe}\SupportedTypes"; ValueType: string; ValueName: ".xlsx"; ValueData: ""; Flags: uninsdeletekey; Tasks: abrircon

[Run]
Filename: "{app}\{#MyAppExe}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall skipifsilent

; Los datos del usuario (%APPDATA%\InventarioPC) NO se borran al desinstalar.
