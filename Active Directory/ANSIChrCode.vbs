dim objWSH, strCHR, c, i, x

set objWSH = createobject("wscript.shell")

choice()

sub choice()
    wscript.echo vbnewline & ucase("select a display method. leave blank to quit")
    wscript.echo vbtab & ucase("1). single character per line")
    wscript.echo vbtab & ucase("2). as an entire string")
    wscript.echo vbtab & ucase("3). as url encode") & vbnewline
    c = wscript.stdin.readline
    if c <> vbnullstring then
      if c = "1" then
        getCHR("1")
      elseif c = "2" then
        getCHR("2")
      elseif c = "3" then
        getCHR("3")
      else
        wscript.echo vbnewline & ucase("invalid choice! please select again or leave blank to quit")
        wscript.sleep 2000
        choice()
      end if
    end if
end sub

sub getCHR(proc)
  i = "1"
  do while i <> vbnullstring
    wscript.echo vbnewline & ucase("enter character(s) to get ansi character code. leave blank to end")
    wscript.echo ucase("for special cases such as 'enter' or 'null' type '{enter}' or '{null}'")
    wscript.echo ucase("only type one character is url encode is selected as display method")
    wscript.echo ucase("type '{change}' to change the display method") & vbnewline
    i = wscript.stdin.readline
    if len(i) = 1 then
      if proc = "1" then
        wscript.echo vbnewline & vbtab & "ANSI code for " & i & " is: chr(" & asc(i) & ")"
      elseif proc = "3" then
        wscript.echo vbnewline & vbtab & "URL code for " & i & " is: " & encURL(i)
      end if
      wscript.sleep 1000
    else
      if proc = "3" then
        wscript.echo vbnewline & vbtab & ucase("please only use one character for url encode display method")
        wscript.sleep 1000
      end if
      if i = "{NULL}" then
        wscript.echo vbnewline & vbtab & "ANSI code for {NULL} is: chr(0)"
      elseif i = "{ENTER}" then
        wscript.echo vbnewline & vbtab & "ANSI code for {ENTER} is: chr(13)"
      elseif i = "{CHANGE}" then
        choice()
      else
        if proc = "1" then
          for x = 1 to len(i)
            wscript.echo vbnewline & vbtab & "ANSI code for " & mid(i, x, 1) & " is: chr(" & asc(mid(i, x, 1)) & ")"
            wscript.sleep 1000
          next
        elseif c = "2" then
          for x = 1 to len(i)
            if x < len(i) then
              strCHR = strCHR & "chr(" & asc(mid(i, x, 1)) & ") & "
            elseif x = len(i) then
              strCHR = strCHR & "chr(" & asc(mid(i, x, 1)) & ")"
            end if
          next
          wscript.echo vbnewline & vbtab & strCHR
          strCHR = vbnullstring
        end if
      end if
    end if
  loop
  wscript.quit
end sub

function encURL(strCHR)
  if strCHR = " " then
    encURL = "%20"
  elseif strCHR = "!" then
    encURL = "%21"
  elseif strCHR = chr(34) then
    encURL = "%22"
  elseif strCHR = "#" then
    encURL = "%23"
  elseif strCHR = "$" then
    encURL = "%24"
  elseif strCHR = "%" then
    encURL = "%25"
  elseif strCHR = "&" then
    encURL = "%26"
  elseif strCHR = "'" then
    encURL = "%27"
  elseif strCHR = "(" then
    encURL = "%28"
  elseif strCHR = ")" then
    encURL = "%29"
  elseif strCHR = "*" then
    encURL = "%2A"
  elseif strCHR = "+" then
    encURL = "%2B"
  elseif strCHR = "," then
    encURL = "%2C"
  elseif strCHR = "-" then
    encURL = "%2D"
  elseif strCHR = "." then
    encURL = "%2E"
  elseif strCHR = "/" then
    encURL = "%2F"
  elseif strCHR = ":" then
    encURL = "%3A"
  elseif strCHR = ";" then
    encURL = "%3B"
  elseif strCHR = "<" then
    encURL = "%3C"
  elseif strCHR = "=" then
    encURL = "%3D"
  elseif strCHR = ">" then
    encURL = "%3E"
  elseif strCHR = "\" then
    encURL = "%5C"
  end if
end function