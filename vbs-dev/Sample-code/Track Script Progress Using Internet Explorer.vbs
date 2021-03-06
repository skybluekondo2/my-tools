Set objExplorer = WScript.CreateObject("InternetExplorer.Application")
objExplorer.Navigate "about:blank"   
objExplorer.ToolBar = 0
objExplorer.StatusBar = 0
objExplorer.Width=400
objExplorer.Height = 200 
objExplorer.Left = 0
objExplorer.Top = 0

Do While (objExplorer.Busy)
    Wscript.Sleep 200
Loop    

objExplorer.Visible = 1             
objExplorer.Document.Body.InnerHTML = "Retrieving service information. " _
    & "This might take several minutes to complete."

strComputer = "."
Set colServices = GetObject("winmgmts:\\" & strComputer & "\root\cimv2"). _
    ExecQuery("Select * from Win32_Service")
For Each objService in colServices
    Wscript.Sleep 200
Next

objExplorer.Document.Body.InnerHTML = "Service information retrieved."
Wscript.Sleep 3000
objExplorer.Quit

