Option Explicit
Dim se, fso
Dim argv, file, str
'VBScript Encoder
'Author: Demon
'Website: http://demon.tw
Set se  = CreateObject("Scripting.Encoder")
Set fso = CreateObject("Scripting.FilesystemObject")
'For Each argv In WScript.Arguments
argv = "C:\Temp\vbscripts\getOSversion.vbs"
    Set file = fso.OpenTextFile(argv)
    str = file.ReadAll
    file.Close
    str = se.EncodeScriptFile(".vbs", str, 0 , "")
    argv = Left(argv, Len(argv)-3) & ".vbe"
    Set file = fso.OpenTextFile(argv, 2, True)
    file.Write str
    file.Close
'Next
MsgBox "OK", vbInformation