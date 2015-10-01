Option Explicit

Dim mainWindow, tableRows, tableFrame, iconsFrame, blacklist

blacklist = Array("Act -","H.C.S. -","Notice -","Transaction","Will Registered")

Function Window(url)
  Dim objShell, objShellWindows, objIE, strURL, i, win

  Set objShell = CreateObject("Shell.Application")
  Set objShellWindows = objShell.Windows
  If objShellWindows.Count = 0 Then
    Wscript.Echo "No browser windows are open to the Script Center."
    Wscript.Quit
  End If
  For i = 0 to objShellWindows.Count - 1
    Set objIE = objShellWindows.Item(i)
    strURL = objIE.LocationURL
    If InStr(strURL, url)Then
      Set win = objIE
    End If
  Next
  Set Window = win
End Function

Function Frame(window, fNum)
  Dim iframes, i

  Set iframes = window.document.getElementsByTagName("frame")
  MsgBox "There are " & Count(iframes) & " frames"
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

Sub clickPrintGif
  Dim a

  For Each a In iconsFrame.contentwindow.document.getElementsbyTagname("a")
    If a.GetAttribute("alt") = "Print" Then
      a.click
      Exit For
    End If
  Next
End Sub

Sub clickPrintButton

iconsFrame.contentwindow.document.getElementsByName("cmd_Print").Item(0).Click
End Sub

Sub ProcessLink(link)
  link.click
  'ClickPrintGif()
  'ClickPrintButton() 'opens new window
end Sub

Function Source(window)
  Source = window.document.body.innerHTML
End Function

Sub Save(text, fileName)
  Dim objFSO, outFile, objFile

  Set objFSO = CreateObject("Scripting.FileSystemObject")
  outFile="D:\" & filename & ".html"
  Set objFile = objFSO.CreateTextFile(outFile,True)
  objFile.Write text & vbCrLf
  objFile.Close
End Sub

Sub SaveSource(fileNumber)
  Dim child_window, html

  Set child_window = window("http://pride/pride/Search/PrintDetails.aspx")
  html = Source(child_window)
  MsgBox html
  Save html, fileNumber
  child_window.Quit
End Sub

Sub Main
  Dim links, num_links, link, counter, row, date, dateStr, dayNum, docType,
blackType, black, found

  Set links = tableFrame.contentwindow.document.getElementsbyTagname("a")
  num_links = Count(links)
  MsgBox num_links & " links in this frame"
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
        'ProcessLink(link)
        'SaveSource(counter)
      end if
    end if
    counter = counter + 1
    if counter = 1 Then
      Exit Sub
    end if
  Next
  'do while mainWindow.Busy
  'loop
  MsgBox found
End Sub



'Set mainWindow =
window("http://www.quackit.com/html/templates/frames/frames_example_6.html")
Set mainWindow = window("http://pride/pride/Main/default.aspx")
'MsgBox "main window is var type: " & VarType(mainWindow)

Set tableFrame = mainWindow.document.getElementsByTagName("frame").Item(0)
Set tableRows =
tableFrame.contentwindow.document.getElementsbyTagname("table").Item(0).rows
Set iconsFrame = mainWindow.document.getElementsByTagName("frame").Item(2)
'IS PRINT BUTTON IN DIFFERENT FRAME - IF SO DEFINE IN clickPrintButton
Main()
