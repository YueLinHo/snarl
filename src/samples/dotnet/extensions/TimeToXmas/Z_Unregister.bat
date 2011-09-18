SET REGASM=c:\windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe
SET DLLFILE=\full phat\snarl\extensions\TimeToXmas\TimeToXmas.dll

"%REGASM%" "%APPDATA%%DLLFILE%" /unregister

PAUSE
