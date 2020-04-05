on error resume next

Dim objWSH, objXCL, objWB, objWS, strEntry, arrCSD, cX, rwX
Set objWSH = CreateObject("wscript.shell")
Set objXCL = CreateObject("Excel.application")

wscript.echo "*** Excel.vbs By : Khristos ***"
wscript.sleep 500
wscript.echo "*** Column Headers Will be Set Up First ***"
wscript.sleep 500
wscript.echo "*** Individual Data Will Follow ***"
wscript.sleep 500
wscript.echo "*** Enter Data in a Comma Delimited Format ***"
wscript.sleep 500
wscript.echo "*** I.E. cell1data,cell2data,cell3data... ***"
wscript.sleep 500
wscript.echo "*** Type '{HELP}' For A List of Commands ***"
wscript.sleep 500
wscript.echo "*** Type '{ABORT}' to Exit Script ***"
wscript.sleep 500
rwX = 1
wscript.sleep 1000

objXCL.visible = True
set objWB = objXCL.workbooks.add
set objWS = objWB.sheets("Sheet1")

objWSH.appactivate "cscript.exe"
wscript.sleep 500
wscript.echo "*** Enter Column Headers - Commands Cannot Be Used ***"
strEntry = wscript.stdin.readline

While InStr(1, strEntry, "{") <> 0
	wscript.echo "*** Invalid Input - Commands Cannot Be Used ***"
	wscript.echo "*** Enter Column Headers - Commands Cannot Be Used ***"
	strEntry = wscript.stdin.readline
	wscript.sleep 500
Wend

arrCSD = split(strEntry, ",")

For cX = 0 To UBound(arrCSD)
  objWS.cells(rwX,cX + 1) = arrCSD(cX)
  objWS.cells(rwX,cX + 1).select
  objWS.selection.font.bold = True
  wscript.sleep 250
Next

strEntry = vbnullstring

While strEntry = vbnullstring
  rwX = rwX + 1
  wscript.echo "*** Position At Row " & rwX & " ***"
  wscript.echo "*** Ready For Data Input - Commands Can Now Be Used ***"
  strEntry = wscript.stdin.readline
  wscript.sleep 500
  If strEntry = "{help}" Then
    rwX = rwX - 1
    wscript.echo "*** To Change A Single Cell Type '{c},row#,col#' ***"
     
    wscript.echo "*** To Change An Entire Row to A Specific Column ***"
      wscript.echo vbtab & "*** Type '{c},row#,{col#}{col#}' ***"
      wscript.echo vbtab & "*** Excel.vbs Will Request a Value For Each Cell ***"
        
    wscript.echo "*** To Change An Entire Column to A Specific Row ***"
      wscript.echo vbtab & "*** Type '{c},{row#}{row#},col#' ***"
      wscript.echo vbtab & "*** Excel.vbs Will Request a Value For Each Cell ***"
        
    wscript.echo "*** To Select A Single Cell for Formatting Type '{f},row#,col#' ***"
      
    wscript.echo "*** To Select A Single Row to A Specific Column for Formatting ***"
      wscript.echo vbtab & "*** Type '{f},row#,{col#}{col#}' ***"
      wscript.echo vbtab & "*** Excel.vbs Will Request The Type of Formatting ***"
        
    wscript.echo "*** To Select A Single Column to A Specific Row for Formatting ***"
      wscript.echo vbtab & "*** Type '{f},{row#}{row#},col#' ***"
      wscript.echo vbtab & "*** Excel.vbs Will Request The Type of Formatting ***"
        
  ElseIf strEntry = "{abort}" Then
    wscript.echo "*** Quitting Script - Clearing Values ***"
    wscript.sleep 1000
    Call scrClean()
  Else
    arrCSD = split(strEntry, ",")
    If arrCSD(0) = "{c}" Then
    	rwX = rwX - 1
    	If InStr(1, arrCSD(1), "{") = 0 And InStr(1, arrCSD(2), "{") = 0 Then
          wscript.echo "*** Enter New Data For Cell (" & arrCSD(1) & "," & arrCSD(2) & ") ***"
          strEntry = wscript.stdin.readline
          objWS.cells(arrCSD(1),arrCSD(2)) = strEntry
          if err.number <> 0 then
            wscript.echo err.number & vbtab & err.description
            wscript.sleep 2000
            err.clear
          end if
          wscript.sleep 500
    	ElseIf InStr(1, arrCSD(1), "{") = 0 And InStr(1, arrCSD(2), "{") Then
  	  Dim arrCol
          arrCol = split(arrCSD(2), "}")
          For cX = Mid(arrCol(0), 2, Len(arrCol(0)) - 1) To Mid(arrCol(1), 2, Len(arrCol(1)) - 1)
            wscript.echo "*** Enter New Data For Cell (" & arrCSD(1) & "," & cX & ") ***"
            strEntry = wscript.stdin.readline
            objWS.cells(arrCSD(1),cX) = strEntry
            wscript.sleep 500
          Next
        ElseIf InStr(1, arrCSD(1), "{") And InStr(1, arrCSD(2), "{") = 0 Then
          Dim arrRow
          arrRow = split(arrCSD(1), "}")
          For cX = Mid(arrRow(0), 2, Len(arrRow(0)) - 1) To Mid(arrRow(1), 2, Len(arrRow(1)) - 1)
            wscript.echo "*** Enter New Data For Cell (" & cX & "," & arrCSD(2) & ") ***"
            strEntry = wscript.stdin.readline
            objWS.cells(cX,arrCSD(2)) = strEntry
            if err.number <> 0 then
              wscript.echo err.number & vbtab & err.description
              wscript.sleep 2000
              err.clear
            end if
            wscript.sleep 500
          Next
  	 End If 	
    ElseIf arrCSD(0) = "{f}" Then
    	rwX = rwX - 1
  	  If InStr(1, arrCSD(1), "{") = 0 And InStr(1, arrCSD(2), "{") = 0 Then
  	    'Call fCell(arrCSD(1), arrCSD(2))
  	  ElseIf InStr(1, arrCSD(1), "{") And InStr(1, arrCSD(2), "{") Then
  	    'Call fRC(arrCSD(1), arrCSD(2))
  	  ElseIf InStr(1, arrCSD(1), "{") And InStr(1, arrCSD(2), "{") = 0 Then
  	    'Call fCR(arrCSD(1), arrCSD(2))
  	  End If
    Else
      For cX = 0 To UBound(arrCSD)
        objWS.cells(rwX,cX + 1) = arrCSD(cX)
        wscript.sleep 500
      Next
    End If
  End If
  Set arrCSD = Nothing
  strEntry = vbnullstring
  wscript.sleep 500
Wend

Sub scrClean()
  set objWS = nothing
  set objWB = nothing
  Set objXCL = Nothing
  Set objWSH = Nothing
  wscript.quit
End Sub

''V--V Excel.vbs By : Khristos V--V''
''V--V VampyreMan Studios Inc V--V''
''V--V www.vampryemanstudios.com V--V''