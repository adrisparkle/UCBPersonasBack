@ECHO OFF

set rootpath=%~dp0
set destination="C:\inetpub\wwwroot\RRHH"

mkdir "%destination%\Areas"
robocopy "%rootpath%\Areas" "%destination%\Areas" /E /COPYALL

mkdir "%destination%\bin"
robocopy "%rootpath%\bin" "%destination%\bin" /E /COPYALL

mkdir "%destination%\Content"
robocopy "%rootpath%\Content" "%destination%\Content" /E /COPYALL

mkdir "%destination%\fonts"
robocopy "%rootpath%\fonts" "%destination%\fonts" /E /COPYALL

mkdir "%destination%\Images"
robocopy "%rootpath%\Images" "%destination%\Images" /E /COPYALL

mkdir "%destination%\Scripts"
robocopy "%rootpath%\Scripts" "%destination%\Scripts" /E /COPYALL

mkdir "%destination%\Static"
robocopy "%rootpath%\Static" "%destination%\Static" /E /COPYALL

mkdir "%destination%\Views"
robocopy "%rootpath%\Views" "%destination%\Views" /E /COPYALL

robocopy "%rootpath%\" "%destination%\\" favicon.ico /COPYALL
robocopy "%rootpath%\" "%destination%\\" Global.asax /COPYALL
robocopy "%rootpath%\" "%destination%\\" "packages.config" /COPYALL
robocopy "%rootpath%\" "%destination%\\" "Web.config" /COPYALL

ECHO ON