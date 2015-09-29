Option Explicit

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

Sub ProcessLink(link)
  link.click
  'TBC
end Sub

Function Main
  Dim pride, main_window, links, num_links, link, counter

  pride = "http://10.0.2.2:8888/pride.html"
  Set main_window = Window(pride)
  'MsgBox "main window is var type: " & VarType(main_window)
  Set links = main_window.document.getElementsbyTagname("a")
  num_links = Count(links)
  MsgBox num_links & " links on page"
  counter = 0
  For Each link In links
    ProcessLink(link)
    counter = counter + 1
    if counter = 2 Then
      Exit Function
    end if
  Next
  do while main_window.Busy
  loop
End Function


Main()
