Dim IE

Sub GetIE(URL)
  Dim objInstances, objIE
  Set objInstances = CreateObject("Shell.Application").windows
  If objInstances.Count > 0 Then
    For Each objIE In objInstances
     If InStr(objIE.LocationURL,URL) > 0 then
       Set IE = objIE
     End if
    Next
  End if
End Sub

Sub ClickLink(linknum)
  Dim anchors
  Dim ItemNr

  Set anchors = IE.document.getElementsbyTagname("a")
  anchors.Item(linknum).click
  do while IE.Busy
  loop
End Sub

'This function extracts text from a specific tag by name and index
'e.g. TABLE,0 (1st Table element) or P,1 (2nd Paragraph element)
'set all to 1 to extract all HTML, 0 for only inside text without HTML
Function ExtractTag(TagName,Num,all)
  dim t

  set t = IE.document.getElementsbyTagname(Tagname)
  if all=1 then
    ExtractTag = t.Item(Num).outerHTML
  else
    ExtractTag = t.Item(Num).innerText
  end if
End Function

'test
'GetIE("http://10.0.2.2:8888/pride.html")

'live
GetIE("http://pride/pride/Main/default.aspx")

ClickLink(0)

'ExtractTag("td",5,0),tagText
'MessageModal>tagText
