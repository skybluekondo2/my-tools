rem //------------------------------------------------------
rem //-- FuncTree Execute Batch file
rem //------------------------------------------------------

rem //***********************
rem //* ���ݒ�
rem //***********************
:���ݒ�
set BIN_DIR=C:\home\kondo\bin\FuncTree\win32
set CMD=%BIN_DIR%\functree.exe
set DEST_DIR=
set OPT=-E%BIN_DIR%\func-ansic.tbl

:��������
setlocal enableextensions

:���C������

if "%1"=="" goto �I������

set FP=%~dp1
set FNM=%~n1
set EXT=%~x1
set OUTF=%FP%\ftree-%FNM%.txt

%CMD% %OPT% %1 > %OUTF%

shift

goto ���C������

:�I������
echo "����I���B"

pause

endlocal
