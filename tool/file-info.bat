@echo off
rem //-------------------------------------------------------//
rem // �t�@�C����: file-info.bat
rem // �@�\�T�v	: 
rem // ����		: file-info.bat [base-dir]
rem //			  base-dir - �J�����g�f�B���N�g��
rem //
rem //-------------------------------------------------------//

setlocal enableextensions

rem /############################################################/
rem /# ���ݒ�
rem /############################################################/

set L_DATE=%DATE:~0,4%%DATE:~5,2%%DATE:~8,2%
set L_FILENAME=�t�@�C�����_%~n1_%L_DATE%.txt
set L_BASE_DIR=%~dpn1

rem ##################################//
rem # �t�@�C���o�̓p�X 
rem ##################################//
set G_FILEOUT_PATH=d:\file-info\

if not exist "%G_FILEOUT_PATH%" (
	mkdir %L_FILEOUT_PATH%
)

set L_OUTFILE=%G_FILEOUT_PATH%%L_FILENAME%

echo #######################################################
echo # �o�̓t�@�C��=%L_OUTFILE%
echo #######################################################
pause
rem //-- CSV�t�@�C���̃^�C�g���s�쐬
echo �t�@�C����,�p�X,����,���t,�T�C�Y,�g���q > %L_OUTFILE%

rem //-- �S�t�@�C������
for /r %L_BASE_DIR% %%x in (*.*) do (
echo %%~nxx,%%~dpx,%%~ax,%%~tx,%%~zx,%%~xx >> %L_OUTFILE%
)

endlocal
