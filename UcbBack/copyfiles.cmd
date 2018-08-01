@ECHO OFF

set rootpath=%~dp0
set destination="C:\Users\Adrian\Desktop\www"

mkdir "%destination%\Areas"
robocopy "%rootpath%\Areas" "%destination%\Areas" /E

mkdir "%destination%\bin"
robocopy "%rootpath%\bin" "%destination%\bin" /E

mkdir "%destination%\Content"
robocopy "%rootpath%\Content" "%destination%\Content" /E

mkdir "%destination%\fonts"
robocopy "%rootpath%\fonts" "%destination%\fonts" /E

mkdir "%destination%\Images"
robocopy "%rootpath%\Images" "%destination%\Images" /E

mkdir "%destination%\Scripts"
robocopy "%rootpath%\Scripts" "%destination%\Scripts" /E

mkdir "%destination%\Static"
robocopy "%rootpath%\Static" "%destination%\Static" /E

mkdir "%destination%\Views"
robocopy "%rootpath%\Views" "%destination%\Views" /E

robocopy "C:\Program Files (x86)\Jenkins\workspace\Windows Service Deployment\UcbBack" "C:\Users\Adrian\Desktop\www" favicon.ico
robocopy "%rootpath%\" %destination% Global.asax
robocopy "%rootpath%\" "%destination%\\" "packages.config"
robocopy "%rootpath%\" "%destination%\\" "Web.config"

ECHO ON