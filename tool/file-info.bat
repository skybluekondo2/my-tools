@echo off
rem //-------------------------------------------------------//
rem // ファイル名: file-info.bat
rem // 機能概要	: 
rem // 書式		: file-info.bat [base-dir]
rem //			  base-dir - カレントディレクトリ
rem //
rem //-------------------------------------------------------//

setlocal enableextensions

rem /############################################################/
rem /# 環境設定
rem /############################################################/

set L_DATE=%DATE:~0,4%%DATE:~5,2%%DATE:~8,2%
set L_FILENAME=ファイル情報_%~n1_%L_DATE%.txt
set L_BASE_DIR=%~dpn1

rem ##################################//
rem # ファイル出力パス 
rem ##################################//
set G_FILEOUT_PATH=d:\file-info\

if not exist "%G_FILEOUT_PATH%" (
	mkdir %L_FILEOUT_PATH%
)

set L_OUTFILE=%G_FILEOUT_PATH%%L_FILENAME%

echo #######################################################
echo # 出力ファイル=%L_OUTFILE%
echo #######################################################
pause
rem //-- CSVファイルのタイトル行作成
echo ファイル名,パス,属性,日付,サイズ,拡張子 > %L_OUTFILE%

rem //-- 全ファイル検索
for /r %L_BASE_DIR% %%x in (*.*) do (
echo %%~nxx,%%~dpx,%%~ax,%%~tx,%%~zx,%%~xx >> %L_OUTFILE%
)

endlocal
