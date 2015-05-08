@SET CONFIGDIR="C:\Program Files\Common Files\Microsoft Shared\web server extensions\14\CONFIG"
@SET GACUTIL="C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\bin\gacutil.exe"

IF %2 == ReleaseFoundation goto Foundation
IF %2 == DebugFoundation goto Foundation

rem copy /y CONFIG\stsadmcommands.moss.lapointe.xml %CONFIGDIR%

:Foundation
rem copy /y CONFIG\stsadmcommands.foundation.lapointe.xml %CONFIGDIR%

rem copy /y POWERSHELL\Registration\*.xml %CONFIGDIR%\POWERSHELL\Registration

%GACUTIL% -if %1

