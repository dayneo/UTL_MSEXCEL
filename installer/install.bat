@ECHO OFF

set /P DEST_SCHEMA=destination schema (username/password@tnsname): 

REM -------------------------------------------------------------------------------------------
REM orace install
REM -------------------------------------------------------------------------------------------
for %%j in (*.jar) do cmd /c loadjavarunner.bat %DEST_SCHEMA% %%j
sqlplus %DEST_SCHEMA% @install.sql
