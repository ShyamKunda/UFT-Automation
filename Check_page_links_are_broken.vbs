
Dim objAllLinks, objDesc, objTargetPage
Set objXMLHttp = CreateObject("MSXML2. XmlHttp")
set objTargetPage = Browser(" ").page(" ") 'Modify according to your page
set objDesc = Description.create
objDesc("html tag").value = "a|A"
objDesc("html tag").regularexpression = true
set objAllLinks = objTargetPage.childobjects(objDesc)

For i = 0 to oAllLinks.count-1 
  If oAllLinks( i). Exist( 0) Then
    On error resume next 
    sHref = oAllLinks( i).Object.href 
    j = j + 1 
    print j & ": " & sHref 
    call checkLink( sHref) 
    if If err.number < > 0 Then 
      reporter.ReportEvent micWarning, "Check Link", "Error: " & err.number & " - Description: " & err.description 
    End If
    On error goto 0 
  else 
    reporter.ReportEvent micWarning, "Check Link", oAllLinks( i). GetTOProperty(" href") & " does not exist." 
  End If 
Next 
print "Total number of processed links: " & j disposeXMLHttp()

Function checkLink( URL) 
    if oXMLHttp.open(" GET", URL, false) = 0 then 
      On error resume next 
      oXMLHttp.send() 
      If oXMLHttp.Status < > 200 Then 
        reporter.ReportEvent micFail, "Check Link", "Link " & URL & " is broken: " & oXMLHttp.Status 
      Else 
        reporter.ReportEvent micPass, "Check Link", "Link " & URL & " is OK" 
      End If 
    End if 
 End Function
 
 Set oXMLHttp = Nothing








