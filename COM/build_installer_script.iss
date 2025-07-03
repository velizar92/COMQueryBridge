#define ApplicationName "COMSQLQueryBridge"

#ifndef ApplicationVersion
#define ApplicationVersion "1.0.0"
#endif

[Files]
Source: "COMQueryBridge\bin\Release\net8.0\*.*"; DestDir: "{app}"

[Setup]
AppName={#ApplicationName}
AppVersion={#ApplicationVersion}
DefaultDirName={pf32}\COMAddins\COMQueryBridge
OutputDir=bin\Setup
OutputBaseFilename={#ApplicationName} Setup {#ApplicationVersion}
VersionInfoProductVersion={#ApplicationVersion}
VersionInfoVersion={#ApplicationVersion}

[Run]
Filename: "regsvr32.exe"; Parameters: "/s ""{app}\COMQueryBridge.dll"""; Flags: runhidden waituntilterminated

[UninstallRun]
Filename: "regsvr32.exe"; Parameters: "/s /u ""{app}\COMQueryBridge.dll"""; Flags: runhidden waituntilterminated



