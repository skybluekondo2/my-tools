Set objDialog = CreateObject("SAFRCFileDlg.FileSave")

objDialog.FileName = "C:\Scripts\Script1.vbs"
objDialog.FileType = "VBScript Script"
intReturn = objDialog.OpenFileSaveDlg

If intReturn Then
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(objDialog.FileName)
    objFile.WriteLine Date
    objFile.Close
Else
    Wscript.Quit
End If

