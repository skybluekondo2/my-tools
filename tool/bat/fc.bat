rem --@echo off

set L_OUTFILE=%~dp1fc_result.txt

rem --setlocal enableextensions

fc %1 %2 > %L_OUTFILE%

rem --endlocal
