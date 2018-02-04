Option Explicit

Private Sub createBook(byval sFile)

		Dim objExcel
		Dim objWkb

		Set objExcel = CreateObject("Excel.Application")

		objExcel.Visible = True
		objExcel.DisplayAlerts = False

		Set objWkb = objExcel.Workbooks.Add
		
		objWkb.SaveAs sFile

		objExcel.Quit

		Set objWkb = Nothing
		Set objExcel =  Nothing
End Sub

Dim objShell
Dim strDesktop

Set objShell = CreateObject("WScript.Shell")

strDesktop = objShell.SpecialFolders("Desktop")

createBook(strDesktop & "/newBook.xls")

Set objShell = Nothing
