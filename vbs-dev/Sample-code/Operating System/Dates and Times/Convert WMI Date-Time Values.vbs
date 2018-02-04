strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objOS = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
 
For Each strOS in objOS
    dtmInstallDate = strOS.InstallDate
    strReturn = WMIDateStringToDate(dtmInstallDate)
    Wscript.Echo strReturn 
Next
 
Function WMIDateStringToDate(dtmInstallDate)
    WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & _
        Mid(dtmInstallDate, 7, 2) & "/" & Left(dtmInstallDate, 4) _
            & " " & Mid (dtmInstallDate, 9, 2) & ":" & _
                Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, _
                    13, 2))
End Function

