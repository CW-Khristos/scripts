dim objWSH, objIE, objFSO, objELEM, strELEM

set objWSH = createobject("wscript.shell")
set objIE = createobject("internetexplorer.application")
set objFSO = createobject("scripting.filesystemobject")

objIE.visible = true
objIE.navigate "http://ilmcw.dyndns.biz"
while (objIE.busy or objIE.readystate <> 4)
  wscript.sleep 100
wend

''LOGIN
wscript.echo objIE.locationurl
if (objIE.locationurl = "http://ilmcw.dyndns.biz/") then
  wscript.echo "Username"
  set objELEM = getElement(objIE.document, "input","usernameid")
  if not (objELEM is nothing) then
    objELEM.value = "cj@thecomputerwarriors.com"
  end if

  set objELEM = nothing
  wscript.echo "Password"
  set objELEM = getElement(objIE.document, "input","passwordfieldid")
  if not (objELEM is nothing) then
    objELEM.value = "Cj062916!"
    'objELEM.text = "Cj062916!"
  end if
  wscript.sleep 5000
  set objELEM = nothing
  wscript.echo "Submit"
  set objELEM = getElement(objIE.document, "form","loginform")
  if not (objELEM is nothing) then
    objELEM.submit
  end if
end if
while (objIE.busy or objIE.readystate <> 4)
  wscript.sleep 100
wend
''ACTIVE ISSUES
'objIE.navigate "http://ilmcw.dyndns.biz/loadActiveIncidents.action"
'while (objIE.busy or objIE.readystate <> 4)
'  wscript.sleep 100
'wend
'if (lcase(objIE.document.title) = "active issues") then
'  wscript.echo "READING ACTIVE ISSUES"
'  set objELEM = getElement(objIE.document, "div","")
'  if not (objELEM is nothing) then
'    if (instr(1, objELEM.id, "activeIncidentGrid-row")) then
'      set colTD = objELEM.getElementsByTagName("td")
'      for each objTD in colTD
'        wscript.echo objTD.class & vbtab & objTD.innertext
'      next
'    end if
'  end if
'end if

wscript.sleep 5000
wscript.quit

function getElement(objParent, elemType, elemID)
  dim colFRM, objFRM

  set getElement = nothing
  set colFRM = objParent.getElementsByTagName(ucase(elemType))

  for each objFRM in colFRM
    wscript.echo objFRM.id & vbtab & objFRM.name
    if (ucase(objFRM.id) = ucase(elemID)) then
      set getElement = objFRM
      exit for
    end if
  next
end function