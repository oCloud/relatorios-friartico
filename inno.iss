[Setup]
AppName=FriarticoPonto
AppVersion=1.0
DefaultDirName={pf}\FriarticoPonto
DefaultGroupName=FriarticoPonto
OutputBaseFilename=FriarticoPontoInstall
Compression=lzma
SolidCompression=yes
DisableDirPage=no

[Files]
Source: "dist\FriarticoPonto.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\FriarticoPonto"; Filename: "{app}\FriarticoPonto.exe"

[Run]
Filename: "{app}\FriarticoPonto.exe"; Description: "Launch YourAppName"; Flags: nowait postinstall skipifsilent
