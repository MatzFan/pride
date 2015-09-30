Option Explicit

Dim main_window
'Set main_window = window("http://pride/pride/Default.aspx")
Set main_window = window("http://10.0.2.2:8888/pride.html")
MsgBox "main window is var type: " & VarType(main_window)

Function Window(url)
  Dim win, objInstances, objIE

  Set objInstances = CreateObject("Shell.Application").windows
  If objInstances.Count > 0 Then
    For Each objIE In objInstances
     If InStr(objIE.LocationURL,url) > 0 then
       Set win = objIE
     End if
    Next
  End if
  Set Window = win
End Function

Function Count(links)
  Dim link, c

  c = 0
  For Each link In links
    c = c + 1
  Next
  Count = c
End Function

Sub clickPrintGif
  Dim a

  For Each a In main_window.Document.GetElementsByTagName("a")
    If a.GetAttribute("alt") = "Print" Then
      a.click
      Exit For
    End If
  Next
End Sub

Sub clickPrintButton
  main_window.Document.getElementsByName("cmd_Print").Item(0).Click
End Sub

Sub ProcessLink(link)
  link.click
  clickPrintGif
  clickPrintButton
end Sub

Function Source(window)
  Source = window.document.body.innerHTML
End Function

Sub Save(text, fileName)
  Dim objFSO, outFile

  Set objFSO = CreateObject("Scripting.FileSystemObject")
  outFile="/Users/me/Desktop/" & filename
  Set objFile = objFSO.CreateTextFile(outFile,True)
  objFile.Write text & vbCrLf
  objFile.Close
End Sub

Sub SaveSource(fileNumber)
  Dim child_window, html

  Set child_window = window("http://pride/pride/Search/PrntDetails.aspx")
  html = Source(child_window)
  Save(html, fileNumber)
  child_window.Quit
End Sub

Function Main
  Dim links, num_links, link, counter

  Set links = main_window.document.getElementsbyTagname("a")
  num_links = Count(links)
  MsgBox num_links & " links on page"
  counter = 0
  For Each link In links
    ProcessLink(link)
    'SaveSource(counter)
    counter = counter + 1
    if counter = 1 Then
      Exit Function
    end if
  Next
  do while main_window.Busy
  loop
End Function


Main()
