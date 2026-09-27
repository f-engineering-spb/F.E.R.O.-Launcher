; F-Engineering Launcher v3 — скрипт инсталлятора (Inno Setup 6, Спринт 4).
; Сборка: scripts\build_installer.cmd
; Кодировка: UTF-8 with BOM (кириллица в имени издателя).

[Setup]
AppName=F-Engineering Launcher
AppVersion=3.1.0
AppPublisher=ООО "Ф-Инжиниринг"
AppPublisherURL=https://github.com/f-engineering-spb/f-engineering-launcher
DefaultDirName={localappdata}\FEngineering_Launcher
DefaultGroupName=F-Engineering Launcher
OutputBaseFilename=FEngineering_Launcher_v3_Setup
OutputDir=..\dist_setup
SetupIconFile=..\app\frontend\assets\flauncher.ico
UninstallDisplayIcon={app}\FEngineeringLauncher.exe
Compression=lzma2/ultra64
SolidCompression=yes
; КРИТИЧНО: без прав администратора, чтобы автоапдейтер мог обновлять файлы без UAC!
PrivilegesRequired=lowest
DisableProgramGroupPage=yes

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; Содержимое собранного PyInstaller (из dist\FEngineeringLauncher)
Source: "..\dist\FEngineeringLauncher\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Исходники и скрипты для динамического обновления и работы
Source: "..\app\*"; DestDir: "{app}\app"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\scripts\*"; DestDir: "{app}\scripts"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\version.json"; DestDir: "{app}"; Flags: ignoreversion
; Создаем скелет папок runtime
Source: "..\runtime\manifests\*"; DestDir: "{app}\runtime\manifests"; Flags: ignoreversion recursesubdirs createallsubdirs external skipifsourcedoesntexist

[Dirs]
Name: "{app}\runtime\cache"
Name: "{app}\runtime\logs"

[Icons]
Name: "{group}\F-Engineering Launcher"; Filename: "{app}\FEngineeringLauncher.exe"
Name: "{group}\Удалить F-Engineering Launcher"; Filename: "{uninstallexe}"
Name: "{autodesktop}\F-Engineering Launcher"; Filename: "{app}\FEngineeringLauncher.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\FEngineeringLauncher.exe"; Description: "{cm:LaunchProgram,F-Engineering Launcher}"; Flags: nowait postinstall skipifsilent
