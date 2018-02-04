Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colFiles = objWMIService.ExecQuery _
	("select * from CIM_DataFile Where Drive = 'C:' AND Extension = 'tmp'")
For Each objFile In colFiles
	'//objFile.Delete
	WScript.Echo objFile.Name & "ÇçÌèúÇµÇ‹ÇµÇΩÅB"
Next
