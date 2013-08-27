@ECHO OFF
set source_dir=%~dp0source
set installer_dir=%~dp0installer
set dep_dir=%~dp0dependancies
set build_dir=%~dp0build
del /F /Q "%build_dir%\*"
xcopy "%source_dir%\*.*"    "%build_dir%\" /C /Y /I /R /H /E
xcopy "%installer_dir%\*.*" "%build_dir%\" /C /Y /I /R /H /E
xcopy "%dep_dir%\*.*"       "%build_dir%\" /C /Y /I /R /H /E
pause