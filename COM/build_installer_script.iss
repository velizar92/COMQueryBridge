#define ApplicationName "COMSQLQueryBridge"

#ifndef ApplicationVersion
#define ApplicationVersion "1.0.0"
#endif

[Setup]
AppName={#ApplicationName}
AppVersion={#ApplicationVersion}
DefaultDirName={pf32}\COMAddins\COMQueryBridge
OutputDir=bin\Setup
OutputBaseFilename={#ApplicationName} Setup {#ApplicationVersion}
VersionInfoProductVersion={#ApplicationVersion}
VersionInfoVersion={#ApplicationVersion}

[Files]
; Copy all output files (DLL, PDB, etc.) from your build folder
Source: "COMQueryBridge\bin\Release\net8.0-windows\*.*"; DestDir: "{app}"; Flags: ignoreversion

; Include a sample config file to be renamed/edited by the user/admin
Source: "COMQueryBridge\bin\Release\net8.0-windows\appsettings.example.json"; DestDir: "{app}"; Flags: ignoreversion

[Run]
; Register the COM DLL silently after installation
Filename: "regsvr32.exe"; Parameters: "/s ""{app}\COMQueryBridge.dll"""; Flags: runhidden waituntilterminated

[UninstallRun]
; Unregister the COM DLL silently during uninstall
Filename: "regsvr32.exe"; Parameters: "/s /u ""{app}\COMQueryBridge.dll"""; Flags: runhidden waituntilterminated
