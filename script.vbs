Option Explicit

Dim objWShell, objShell, mainWindow, tableRows, tableFrame, iconsFrame, blacklist

blacklist = Array("Act -","H.C.S. -","Notice -","Transaction","Will Registered")

Function Window(url)
  Dim objShellWindows, objIE, strURL, i, win

  Do
    Set objShellWindows = objShell.Windows
    For i = 0 to objShellWindows.Count - 1
      Set objIE = objShellWindows.Item(i)
      strURL = objIE.LocationURL
      If InStr(strURL, url)Then
        Set win = objIE
        Exit Do
      End If
    Next
    WScript.sleep 100
  Loop
  Set Window = win
End Function

Function Frame(window, fNum)
  Dim iframes, i

  Set iframes = window.document.getElementsByTagName("frame")
  Frame = iframes.Item(fnum)
End Function

Function Count(items)
  Dim item, c

  c = 0
  For Each item In items
    c = c + 1
  Next
  Count = c
End Function

Sub Save(text, fileName)
  Dim objFSO, outFile, objFile

  Set objFSO = CreateObject("Scripting.FileSystemObject")
  outFile="D:\1984\" & filename & ".html"
  Set objFile = objFSO.CreateTextFile(outFile,True)
  objFile.Write text & vbCrLf
  objFile.Close
End Sub

Sub SaveDetailsWindowSource(fileNumber)
  Dim url, childWindow, html

  url = "http://pride/pride/Search/PrintDetails.aspx"
  Set childWindow = window(url)
  html = childWindow.document.body.innerHTML
  Save html, fileNumber
End Sub

Sub ClosePrintDialog
  Dim box

  Do
    box = objWShell.AppActivate("Print")
    If box = True Then
      objWShell.SendKeys "{ESC}"
      Exit Do
    End If
    WScript.Sleep 100
  Loop
end Sub

Sub ClickIcon
  Dim icon

  Do
    Set icon = iconsFrame.contentwindow.document.getElementsbyTagname("a")(6)
    if IsObject(icon) Then
      Exit Do
    end if
    WScript.sleep 100
  Loop
  icon.click
End Sub

Sub ClickButton
  Dim button

  Do
    Set button = iconsFrame.contentwindow.document.getElementsByName("cmd_Print").Item(0)
    if IsObject(button) Then
      Exit Do
    end if
    WScript.sleep 100
  Loop
  button.click
End Sub

Sub Main
  Dim links, num_links, link, counter, row, date, dateStr, dayNum, docType, blackType, black, found, resp

  Set links = tableFrame.contentwindow.document.getElementsbyTagname("a")
  num_links = Count(links)
  counter = 0
  found = 0
  For Each link In links
    Set row = tableRows(counter + 1) '1 more than links because of title row
    date = Split(row.cells(3).innerText, "/")
    dateStr = date(1) & "/" & date(0) & "/" & date(2)
    dayNum = Weekday(CDate(dateStr))
    if dayNum = 6 Then 'Friday
      docType = row.cells(8).innerText
      black = false
      For Each blackType in blacklist
        If Instr(1, docType, blacktype) = 1 Then '= 1 means true :)
          black = true
          Exit for 'blacklist - not interested
        end if
      Next
      if black = false Then
        found = found + 1
        link.click
        ClickIcon()
        ClickButton()
        ClosePrintDialog()
        SaveDetailsWindowSource(counter)
        if counter Mod 100 = 0 Then
          resp = MsgBox(found, 3, "Continue?")
          if resp = 7 Then
            Exit Sub
          end if
        end if
      end if
    end if
    'if counter = 10 Then
      'exit Sub
    'end if
    counter = counter + 1
  Next
  'MsgBox found
End Sub


Set objWShell = CreateObject("Wscript.Shell")
Set objShell = CreateObject("Shell.Application")
Set mainWindow = window("http://pride/pride/Main/default.aspx")
Set tableFrame = mainWindow.document.getElementsByTagName("frame").Item(2)
Set iconsFrame = mainWindow.document.getElementsByTagName("frame").Item(3)
Set tableRows = tableFrame.contentwindow.document.getElementsbyTagname("table").Item(0).rows
Main()
