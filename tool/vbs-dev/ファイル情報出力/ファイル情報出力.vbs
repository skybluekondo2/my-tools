'// �ϐ������I�錾�I�v�V�����w��
Option Explicit

'// �O���[�o���萔 �錾
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
	'//-- �p�����[�^���擾����B
	'//-------------------------------------------------------//
	Set args = WScript.Arguments

	if WScript.Arguments.Count <> 1 then
	   WScript.Echo args.showUsage & "[file]"
	   WScript.Quit 
	end if

	'//-------------------------------------------------------//
	'//-- FileSystemObject�I�u�W�F�N�g�𐶐�
	'//-------------------------------------------------------//
	Set g_objFSO = CreateObject("Scripting.FileSystemObject")

	'//-------------------------------------------------------//
	'//�t�H���_�̑��݃`�F�b�N
	'//-------------------------------------------------------//
	fname=args.Item(0)

	If g_objFSO.FolderExists(fname) Then
	Else
		g_objCSVFile.WriteLine "�t�H���_�����݂܂���I dir=" & fname
		WScript.Quit 1
	End If

	'//-- �t�H���_�I�u�W�F�N�g�𐶐�����B
	Set objFolder = g_objFSO.GetFolder(fname) 

	'//-- �e�L�X�g�t�@�C��(CSV)�I�u�W�F�N�g�𐶐�����B
	strFilePath = "D:\file-info\" & _
				"�t�@�C�����_" & objFolder.Name & "_" & _
				getDateString() & ".txt"
	Set g_objCSVFile = g_objFSO.OpenTextFile(strFilePath, ForWriting, True)

	'// �b�r�u�t�@�C���̃^�C�g���s���쐬����B
	g_objCSVFile.WriteLine "�t�@�C����,�T�C�Y,�쐬��,�ŏI�X�V��,�ŏI�A�N�Z�X,�t���p�X"

	'// �t�@�C�����o�̓v���V�[�W���[���b����������B
	ShowDIR(objFolder)

	'// �����������b�Z�[�W��\������B
	Msgbox "�t�@�C�����o�͂��������܂����B"

	'// �s�v�I�u�W�F�N�g��j������B
	Set g_objCSVFile = Nothing
	Set g_objFSO = Nothing
	Set args = Nothing

End Sub
'//#############################################################################
'//	�v���V�[�W���[��: ShowDIR()
'//	�@�\�T�v		: �t�@�C�������b�r�u�t�@�C���`���ŏo�͂���B
'//	����			: ShowDIR [base-dir]
'//	�p�����[�^		: base-dir - �J�����g�f�B���N�g�����w�肷��B
'//#############################################################################
Sub ShowDIR(byRef p_objFolder)
	Dim file,dispname,F
	Dim g_objFSOd,objSubFld

    '//�t�@�C���R���N�V����
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

	'// �T�u�t�H���_����������B
	Set objSubFld = p_objFolder.SubFolders
	For Each file in objSubFld
	  '//--g_objCSVFile.WriteLine file.Name & "," & file.Attributes
	  ShowDIR(file) 
	Next
End Sub

call main
