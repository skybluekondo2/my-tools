Set objSession = CreateObject("Microsoft.Update.Session")
Set objSearcher = objSession.CreateUpdateSearcher
intCount = objSearcher.GetTotalHistoryCount
If intCount > 0 Then
	Set colHistory = objSearcher.QueryHistory(0, intCount)
	WScript.Echo "Date               Title"
	WScript.Echo "------------------ ----------------------------------------"
	For Each objHistory In colHistory
		WScript.Echo Mid(objHistory.Date, 1, 19) & " " & objHistory.Title
	Next
End If
	