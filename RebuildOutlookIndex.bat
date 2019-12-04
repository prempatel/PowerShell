net stop wsearch
move "%programdata%\microsoft\search\data\applications\windows\Windows.edb" "%programdata%\microsoft\search\data\applications\windows\Windows.edb.bak"
del "%programdata%\microsoft\search\data\applications\windows\Windows.edb"
:wsearch
net start wsearch
echo off
IF NOT %ERRORLEVEL%==0 (goto :wsearch) ELSE goto :END
:END