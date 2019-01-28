@ECHO OFF

set rootpath=%~dp0
set destinationprod="C:\inetpub\wwwroot\RRHH"
set destination="C:\Users\Adrian\Desktop\dev\www2"

call :strLen rootpath strlen
set /a strlen=%strlen%-8

CALL SET prevpath=%%rootpath:~0,%strlen%%%
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe  "C:\Users\Adrian\Desktop\dev\UcbBack - 4.5.2\UcbBack.sln" /p:Configuration=Debug /p:Platform="Any CPU" /p:VisualStudioVersion=12.0 /t:Rebuild

:: rmdir /s /q "%destination%\"

mkdir "%destination%\Areas"
robocopy "%rootpath%\Areas" "%destination%\Areas" /E /COPYALL /it /NFL /NDL /NJH /NJS /nc /ns /np

mkdir "%destination%\bin"
robocopy "%rootpath%\bin" "%destination%\bin" /E /COPYALL /it /NFL /NDL /NJH /NJS /nc /ns /np

mkdir "%destination%\Content"
robocopy "%rootpath%\Content" "%destination%\Content" /E /COPYALL /it /NFL /NDL /NJH /NJS /nc /ns /np

mkdir "%destination%\fonts"
robocopy "%rootpath%\fonts" "%destination%\fonts" /E /COPYALL /it /NFL /NDL /NJH /NJS /nc /ns /np

mkdir "%destination%\Images"
robocopy "%rootpath%\Images" "%destination%\Images" /E /COPYALL /is /NFL /NDL /NJH /NJS /nc /ns /np

mkdir "%destination%\Scripts"
robocopy "%rootpath%\Scripts" "%destination%\Scripts" /E /COPYALL /it /NFL /NDL /NJH /NJS /nc /ns /np

mkdir "%destination%\Views"
robocopy "%rootpath%\Views" "%destination%\Views" /E /COPYALL /it /NFL /NDL /NJH /NJS /nc /ns /np

robocopy "%rootpath%\" "%destination%\\" favicon.ico /COPYALL /it /NFL /NDL /NJH /NJS /nc /ns /np
robocopy "%rootpath%\" "%destination%\\" Global.asax /COPYALL /it /NFL /NDL /NJH /NJS /nc /ns /np
robocopy "%rootpath%\" "%destination%\\" "packages.config" /COPYALL /it /NFL /NDL /NJH /NJS /nc /ns /np
robocopy "%rootpath%\" "%destination%\\" "Web.config" /COPYALL /it /NFL /NDL /NJH /NJS /nc /ns /np




mkdir "%destination%\Static"
robocopy "%prevpath%\Front\dist\static" "%destination%\Static" /E /COPYALL /is /NFL /NDL /NJH /NJS /nc /ns /np
robocopy "%prevpath%\Front\dist\\" "%destination%\Views\Home\\" "index.html" /COPYALL /is /NFL /NDL /NJH /NJS /nc /ns /np


echo "@{    Layout = "";   }" > "%destination%\Views\Home\Index.cshtml"
type "%destination%\Views\Home\index.html" >> "%destination%\Views\Home\Index.cshtml"

ECHO ON
exit /b

:strLen
setlocal enabledelayedexpansion
:strLen_Loop
  if not "!%1:~%len%!"=="" set /A len+=1 & goto :strLen_Loop
(endlocal & set %2=%len%)
goto :eof