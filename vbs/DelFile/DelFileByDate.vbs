strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colFiles = objWMIService.ExecQuery _
	("Select * from CIM_DataFile Where Drive = 'C:'")
For Each objFile In colFiles
	strDate = WMIDateStringToDate(objFile.LastModified)
	If DateDiff("m", strDate, Now) >= 1 Then
		'// objFile.Delete
		WScript.Echo objFile.Name & "(" & strDate & ")ÇçÌèúÇµÇ‹ÇµÇΩÅB"
	End If
Next

Private Function WMIDateStringToDate( _
	byval dtmInstallDate _
)
	WMIDateStringToDate = CDate( _
		Mid(dtmInstallDate, 5, 2) & "/" & _
		Mid(dtmInstallDate, 7, 2) & "/" & _
		Left(dtmInstallDate, 4) & " " & _
		Mid(dtmInstallDate, 9, 2) & ":" & _
		Mid(dtmInstallDate, 11, 2) & ":" & _
		Mid(dtmInstallDate, 13, 2) _
		)
End Function
