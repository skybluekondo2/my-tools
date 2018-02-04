rem //-- 環境設定
set L_DATE=%DATE:~2,4%%DATE:~7,2%%DATE:~10,2%
set L_FILENAME=ファイル一覧.txt
set L_BACKUP=ファイル一覧_%L_DATE%.txt
set L_BASE_DIR=C:\home\kondo\ＡＬＳＯＫ\ＧＣシステム\91.ドキュメント管理
set L_OUTFILE=%L_BASE_DIR%\%L_FILENAME%

rem //-- ファイルバックアップ
ren %L_OUTFILE% %L_BACKUP%

rem //-- CSVファイルのタイトル行作成
echo filename,path,attribute,date,size,ext > %L_OUTFILE%

rem //-- 全ファイル検索
for /r b:\ %%x in (*.*) do (
echo %%~nxx,%%~dpx,%%~ax,%%~tx,%%~zx,%%~xx >> %L_OUTFILE%
)
