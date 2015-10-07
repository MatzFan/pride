Option Explicit

Dim objWShell, objShell, mainWindow, tableRows, tableFrame, iconsFrame, blacklist, year

year = "2015"
blacklist = Array("Act -","Co ownership Declaration","H.C.S. -","Notice -","Rights - Sale","Transaction","Will Registered")

Function Window(url)
  Dim objWindow, win
  do
    for each objWindow in objShell.Windows
      if InStr(objWindow.FullName, "iexplore") then
        if InStr(objWindow.LocationURL, url) then
          Set win = objWindow
          exit do
        end if
      end if
    next
    WScript.sleep 10
  loop
  Set Window = win
End Function
'23
Function Count(items)
  Dim item, c
  c = 0
  for each item in items
    c = c + 1
  next
  Count = c
End Function
'32
Sub Save(text, fileName)
  Dim objFSO, outFile, objFile
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  outFile="D:\" & year & "\" & filename & ".html"
  Set objFile = objFSO.CreateTextFile(outFile, True)
  objFile.Write text & vbCrLf
  objFile.Close
End Sub
'41
Sub SaveDetailsWindowSource(fileNumber)
  Save window("http://pride/pride/Search/PrintDetails.aspx").document.body.innerHTML, fileNumber
End Sub
'45
Sub ClosePrintDialog
  do
    if objWShell.AppActivate("Print") = True then
      objWShell.SendKeys "{ESC}"
    else
      exit do
    end if
    WScript.Sleep 10
  loop
end Sub
'56
Sub ClickIt(object)
  do
    On Error Resume Next
    object.click
    if Err.Number = 0 then
      On Error Goto 0
      Err.Clear
      exit do
    end if
    WScript.sleep 10
  loop
End Sub
'69
Sub ClickIcon
  Dim icon
  do
    Set icon = iconsFrame.contentwindow.document.getElementsbyTagname("a")(6)
    if icon is Nothing then
    else
      exit do
    end if
    WScript.sleep 10
  loop
  ClickIt(icon)
End Sub
'82
Sub ClickButton
  Dim button
  do
    Set button = iconsFrame.contentwindow.document.getElementsByName("cmd_Print").Item(0)
    if button is Nothing then
    else
      exit do
    end if
    WScript.sleep 10
  loop
  ClickIt(button)
End Sub
'95
Sub ClosePrintDetailsWindow
  Dim objWindow, gone
  do
    gone = True
    for each objWindow in objShell.Windows
      if InStr(objWindow.FullName, "iexplore") then
        strURL = objWindow.LocationURL
        if InStr(strURL, "http://pride/pride/Search/PrintDetails.aspx") then
          gone = False
          objWindow.Quit
        end if
      end if
      if gone = True then
        exit do
      end if
    next
    Wscript.sleep 10
  loop
end Sub
'115
Sub Main
  Dim links, num_links, link, counter, row, date, dateStr, dayNum, docType, blackType, black, found, resp

  Set links = tableFrame.contentwindow.document.getElementsbyTagname("a")
  num_links = Count(links)
  counter = 0
  found = 0
  for each link in links
    Set row = tableRows(counter + 1) '1 more than links because of title row
    date = Split(row.cells(3).innerText, "/")
    dateStr = date(1) & "/" & date(0) & "/" & date(2)
    dayNum = Weekday(CDate(dateStr))
    if dayNum = 6 then 'Friday
      docType = row.cells(8).innerText
      black = False
      for each blackType in blacklist
        if Instr(1, docType, blacktype) = 1 then '= 1 means True :)
          black = True
          exit for 'blacklist - not interested
        end if
      next
      if black = False then
        found = found + 1
        link.click
        ClickIcon()
        ClickButton()
        ClosePrintDialog()
        SaveDetailsWindowSource(counter)
        ClosePrintDetailsWindow()
        'if counter > 99 And counter Mod 100 = 0 Then
          'resp = MsgBox(found, 3, "Continue?")
          'if resp = 7 Then
            'Exit Sub
          'end if
        'end if
      end if
    end if
    counter = counter + 1
  next
End Sub

Set objWShell = CreateObject("Wscript.Shell")
Set objShell = CreateObject("Shell.Application")
Set mainWindow = window("http://pride/pride/Main/default.aspx")
Set tableFrame = mainWindow.document.getElementsByTagName("frame").Item(2)
Set iconsFrame = mainWindow.document.getElementsByTagName("frame").Item(3)
Set tableRows = tableFrame.contentwindow.document.getElementsbyTagname("table").Item(0).rows
Main()
MsgBox "Done"
