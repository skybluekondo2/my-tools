Set arg = WScript.Arguments
SearchString = arg(0)
Set objSession = CreateObject("Microsoft.Update.Session")
Set objSearcher = objSession.CreateUpdateSearcher
intCount = objSearcher.GetTotalHistoryCount
If intCount > 0 Then
	Set colHistory = objSearcher.QueryHistory(0, intCount)
	WScript.Echo "Date               Title"
	WScript.Echo "---                -----"
	For Each objHistory In colHistory
		If (objHistory.HResult = 0) AND (InStr(objHistory.Title, SearchString) > 0) Then
			WScript.Echo Mid(objHistory.Date, 1, 19) & " " & objHistory.Title
		End If
	Next
End If
	