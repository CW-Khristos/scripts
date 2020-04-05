Dim strName, strFile, strKey, strTMP, strCR, strCode
Dim objFSO, objFSO1, objCode, objTCode, intCR

WScript.Echo "Enter Script Name:"
strName = WScript.StdIn.ReadLine
WScript.Echo "Enter File:"
strFile = WScript.StdIn.ReadLine
WScript.Echo "Enter Password:"
strKey = WScript.StdIn.ReadLine

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO1 = CreateObject("Scripting.FileSystemObject")
Set objTCode = objFSO1.CreateTextFile(strName & ".vbs", TRUE)
Set objCode = objFSO.OpenTextFile(strFile, 1)

objTCode.WriteLine "On Error Resume Next"
objTCode.WriteLine "Set objWSH = CreateObject(" & chr(34) & "WScript.Shell" & chr(34) & ")"
objTCode.WriteLine "Set objFSO = CreateObject(" & chr(34) & "Scripting.FileSystemObject" & chr(34) & ")"
objTCode.WriteLine "Set objFSO1 = CreateObject(" & chr(34) & "Scripting.FileSystemObject" & chr(34) & ")"
objTCode.WriteLine "Set objTCode = objFSO1.CreateTextFile(" & chr(34) & "tmp.vbs" & chr(34) & ", TRUE)"
objTCode.WriteLine "Set objCode = objFSO.OpenTextFile(" & chr(34) & strName & ".vbs" & chr(34) & ")"
objTCode.WriteLine "Do Until objCode.AtEndOfStream"
objTCode.WriteLine "strTMP = objCode.ReadLine"
objTCode.WriteLine "For x = 1 to Len(strTMP)"
objTCode.WriteLine "strCR = Mid(strTMP, x, 1)"
objTCode.WriteLine "If strCR <> " & chr(34) & "'" & chr(34) & " Then"
objTCode.WriteLine "If Asc(strCR) < 32 Then"
objTCode.WriteLine "intCR = ((Asc(strCR) + 255) / 2) - " & Len(strKey)
objTCode.WriteLine "Else"
objTCode.WriteLine "intCR = (Asc(strCR) / 2) - " & Len(strKey)
objTCode.WriteLine "End If"
objTCode.WriteLine "strCR = chr(intCR)"
objTCode.WriteLine "End If"
objTCode.WriteLine "strCode = strCode & strCR"
objTCode.WriteLine "Next"
objTCode.WriteLine "If Mid(strCode, 1, 1) = " & chr(34) & "'" & chr(34) & " Then"
objTCode.WriteLine "objTCode.WriteLine Mid(strCode, 2, Len(strCode) - 1)"
objTCode.WriteLine "End If"
objTCode.WriteLine "strCode = VBNullString"
objTCode.WriteLine "Loop"
objTCode.WriteLine "objTCode.Close"
objTCode.WriteLine "objWSH.run " & chr(34) & "tmp.vbs" & chr(34)
objTCode.WriteLine "wscript.sleep 1000"
objTCode.WriteLine "objFSO.DeleteFile " & chr(34) & "tmp.vbs" & chr(34) & ", TRUE"


Do Until objCode.AtEndOfStream

strTMP = objCode.ReadLine

For x = 1 to Len(strTMP)
strCR = Mid(strTMP, x, 1)
intCR = (asc(strCR) + Len(strKey)) * 2

If intCR > 255 then
intCR = (intCR - 255)
End If

strCR = chr(intCR)
strCode = strCode & strCR

Next

objTCode.WriteLine "'" & strCode
strCode = VBNullString

Loop

objTCode.Close

Set objTCode = Nothing
Set objCode = Nothing
Set objFSO1 = Nothing
Set objFSO = Nothing