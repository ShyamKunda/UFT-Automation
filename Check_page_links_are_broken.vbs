Dim objAllLinks, objDesc, objTargetPage, objXMLHttp
Set objXMLHttp = CreateObject("MSXML2. XmlHttp")
set objTargetPage = Browser(" ").page(" ") 'Modify according to your page
set objDesc = Description.create
objDesc("html tag").value = "a|A"
objDesc("html tag").regularexpression = true
set objAllLinks = objTargetPage.childobjects(objDesc)

For i = 0 to objAllLinks.count-1 
  If objAllLinks( i).Exist( 0) Then
    On error resume next 
    sHref = objAllLinks( i).Object.href 
    j = j + 1 
    print j & ": " & sHref 
    call checkLink( sHref) 
    if If err.number < > 0 Then 
      reporter.ReportEvent micWarning, "Check Link", "Error: " & err.number & " - Description: " & err.description 
    End If
    On error goto 0 
  else 
    reporter.ReportEvent micWarning, "Check Link", objAllLinks( i).GetTOProperty(" href") & " does not exist." 
  End If 
Next 
print "Total number of processed links: " & j 

Function checkLink( URL) 
    if objXMLHttp.open(" GET", URL, false) = 0 then 
      On error resume next 
      objXMLHttp.send() 
      If objXMLHttp.Status < > 200 Then 
        reporter.ReportEvent micFail, "Check Link", "Link " & URL & " is broken: " & objXMLHttp.Status 
      Else 
        reporter.ReportEvent micPass, "Check Link", "Link " & URL & " is OK" 
      End If 
    End if 
 End Function
 
 Set objXMLHttp = Nothing
