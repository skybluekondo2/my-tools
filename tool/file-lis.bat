rem //-- ���ݒ�
set L_DATE=%DATE:~2,4%%DATE:~7,2%%DATE:~10,2%
set L_FILENAME=�t�@�C���ꗗ.txt
set L_BACKUP=�t�@�C���ꗗ_%L_DATE%.txt
set L_BASE_DIR=C:\home\kondo\�`�k�r�n�j\�f�b�V�X�e��\91.�h�L�������g�Ǘ�
set L_OUTFILE=%L_BASE_DIR%\%L_FILENAME%

rem //-- �t�@�C���o�b�N�A�b�v
ren %L_OUTFILE% %L_BACKUP%

rem //-- CSV�t�@�C���̃^�C�g���s�쐬
echo filename,path,attribute,date,size,ext > %L_OUTFILE%

rem //-- �S�t�@�C������
for /r b:\ %%x in (*.*) do (
echo %%~nxx,%%~dpx,%%~ax,%%~tx,%%~zx,%%~xx >> %L_OUTFILE%
)
