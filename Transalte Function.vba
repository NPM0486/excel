Function Translate(strST As String, strSLC As String, strTLC As String) As String

'1 - url
Dim strURL As String
strURL = "https://translate.google.com/m?sl=" & strSLC & "&tl=" & strTLC & "&hl=en&ie=UTF-8&q=" & strST

'2 - API XML HTTP
Dim XMLHTTP As Object
Set XMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
XMLHTTP.Open "GET", strURL, False
XMLHTTP.setrequestheader "User-Agent", "Mozilla/5.0 (compatible;MSIE 6.0; Windows NT 10.0))"
XMLHTTP.send ""

'3 - HTML file
Dim ObjHTML As Object
Set ObjHTML = CreateObject("HTMLFile")
With ObjHTML
    .Open
    .write XMLHTTP.responseText
    .Close
End With

'4 - read HTML file
'Microsoft HTML Object Library
Dim HTMLDoc As HTMLDocument
Set HTMLDoc = ObjHTML

Dim ObjClass As Object
Set ObjClass = HTMLDoc.getElementsByClassName("result-container")(0)
If Not ObjClass Is Nothing Then
    Translate = ObjClass.innerText
End If

'relasing the memory
Set ObjClass = Nothing
Set ObjHTML = Nothing
Set XMLHTTP = Nothing

End Function