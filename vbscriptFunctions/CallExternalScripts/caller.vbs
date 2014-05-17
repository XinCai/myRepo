' Caller vbs file
Dim fso 
Set fso = createobject("scripting.filesystemobject") 
executeglobal fso.opentextfile("C:\Temp\vbscript\functions.vbs ", 1).readall 
Set fso = Nothing

'start calling external function
msgbox "1+2=" & fun(1,2)