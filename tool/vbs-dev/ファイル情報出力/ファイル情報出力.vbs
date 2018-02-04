'// 変数明示的宣言オプション指定
Option Explicit

'// グローバル定数 宣言
Const Normal = 0
Const Directory = 16
Const Archive = 32
Const ForReading = 1
Const ForWriting = 2

Public g_objFSO
Public g_objCSVFile

Private Function getDateString
	Dim strDate
	Dim strDateString

	strDate = date()

	strDateString = _
			mid(strDate,1,4) & _
			mid(strDate, 6, 2) & _
			mid(strDate, 9, 2)
	getDateString = strDateString
End Function

Private Sub Main
	Dim args
	Dim fname
	Dim objFolder
	Dim strFilePath

	'//-------------------------------------------------------//
	'//-- パラメータを取得する。
	'//-------------------------------------------------------//
	Set args = WScript.Arguments

	if WScript.Arguments.Count <> 1 then
	   WScript.Echo args.showUsage & "[file]"
	   WScript.Quit 
	end if

	'//-------------------------------------------------------//
	'//-- FileSystemObjectオブジェクトを生成
	'//-------------------------------------------------------//
	Set g_objFSO = CreateObject("Scripting.FileSystemObject")

	'//-------------------------------------------------------//
	'//フォルダの存在チェック
	'//-------------------------------------------------------//
	fname=args.Item(0)

	If g_objFSO.FolderExists(fname) Then
	Else
		g_objCSVFile.WriteLine "フォルダが存在ません！ dir=" & fname
		WScript.Quit 1
	End If

	'//-- フォルダオブジェクトを生成する。
	Set objFolder = g_objFSO.GetFolder(fname) 

	'//-- テキストファイル(CSV)オブジェクトを生成する。
	strFilePath = "D:\file-info\" & _
				"ファイル情報_" & objFolder.Name & "_" & _
				getDateString() & ".txt"
	Set g_objCSVFile = g_objFSO.OpenTextFile(strFilePath, ForWriting, True)

	'// ＣＳＶファイルのタイトル行を作成する。
	g_objCSVFile.WriteLine "ファイル名,サイズ,作成日,最終更新日,最終アクセス,フルパス"

	'// ファイル情報出力プロシージャーをＣａｌｌする。
	ShowDIR(objFolder)

	'// 処理完了メッセージを表示する。
	Msgbox "ファイル情報出力が完了しました。"

	'// 不要オブジェクトを破棄する。
	Set g_objCSVFile = Nothing
	Set g_objFSO = Nothing
	Set args = Nothing

End Sub
'//#############################################################################
'//	プロシージャー名: ShowDIR()
'//	機能概要		: ファイル情報をＣＳＶファイル形式で出力する。
'//	書式			: ShowDIR [base-dir]
'//	パラメータ		: base-dir - カレントディレクトリを指定する。
'//#############################################################################
Sub ShowDIR(byRef p_objFolder)
	Dim file,dispname,F
	Dim g_objFSOd,objSubFld

    '//ファイルコレクション
    Set F = p_objFolder.Files  

	For Each file in F
	  '//--g_objCSVFile.WriteLine file.Name & "," & file.Attributes
	  if (file.Attributes = Normal Or  _
	     file.Attributes = Archive) Then
	      g_objCSVFile.WriteLine _
	          file.Name & "," & _
	          file.Size & "," & _
	          file.DateCreated & "," & _
	          file.DateLastModified & "," & _
	          file.DateLastAccessed & "," & _
	          file.Path
	  End If
	Next

	'// サブフォルダを検索する。
	Set objSubFld = p_objFolder.SubFolders
	For Each file in objSubFld
	  '//--g_objCSVFile.WriteLine file.Name & "," & file.Attributes
	  ShowDIR(file) 
	Next
End Sub

call main
