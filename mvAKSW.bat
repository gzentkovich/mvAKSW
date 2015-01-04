@ECHO off
SETLOCAL ENABLEDELAYEDEXPANSION
SET DIR_SRC=C:\inetpub\ftproot
SET DIR_DST=F:\Akcelerant\ASKW
SET DIR_ARC=F:\Archives

FOR /F "tokens=1-3 delims=/-" %%A IN ("%Today%") DO (
    SET MONTH=%%A
    SET DAY=%%B
    SET YEAR=%%C
)
SET CURRDATE=%MONTH%-%DAY%-%YEAR%

ECHO Set objArgs = WScript.Arguments > _zipIt.vbs
ECHO InputFolder = objArgs(0) >> _zipIt.vbs
ECHO ZipFile = objArgs(1) >> _zipIt.vbs
ECHO CreateObject("Scripting.FileSystemObject").CreateTextFile(ZipFile, True).Write "PK" ^& Chr(5) ^& Chr(6) ^& String(18, vbNullChar) >> _zipIt.vbs
ECHO Set objShell = CreateObject("Shell.Application") >> _zipIt.vbs
ECHO Set source = objShell.NameSpace(InputFolder).Items >> _zipIt.vbs
ECHO objShell.NameSpace(ZipFile).CopyHere(source) >> _zipIt.vbs
ECHO wScript.Sleep 2000 >> _zipIt.vbs

CScript _zipIt.vbs %DIR_SRC% %DIR_ARC%\AKSW_%CURRDATE%.zip
MOVE %DIR_SRC%\* %DIR_DST%\
ENDLOCAL
