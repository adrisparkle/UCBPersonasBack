@ECHO OFF

set rootpath=%~dp0
set destination="C:\Users\Adrian\Desktop\www"

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

robocopy "C:\Program Files (x86)\Jenkins\workspace\Windows Service Deployment\UcbBack" "C:\Users\Adrian\Desktop\www" favicon.ico
robocopy "%rootpath%\" %destination% Global.asax
robocopy "%rootpath%\" "%destination%\\" "packages.config"
robocopy "%rootpath%\" "%destination%\\" "Web.config"

ECHO ON