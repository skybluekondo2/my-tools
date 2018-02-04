rem //------------------------------------------------------
rem //-- FuncTree Execute Batch file
rem //------------------------------------------------------

rem //***********************
rem //* 環境設定
rem //***********************
:環境設定
set BIN_DIR=C:\home\kondo\bin\FuncTree\win32
set CMD=%BIN_DIR%\functree.exe
set DEST_DIR=
set OPT=-E%BIN_DIR%\func-ansic.tbl

:初期処理
setlocal enableextensions

:メイン処理

if "%1"=="" goto 終了処理

set FP=%~dp1
set FNM=%~n1
set EXT=%~x1
set OUTF=%FP%\ftree-%FNM%.txt

%CMD% %OPT% %1 > %OUTF%

shift

goto メイン処理

:終了処理
echo "正常終了。"

pause

endlocal
