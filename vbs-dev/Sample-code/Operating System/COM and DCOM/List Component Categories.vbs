On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_ComponentCategory")

For Each objItem in colItems
    Wscript.Echo "Category ID: " & objItem.CategoryId
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo
Next

