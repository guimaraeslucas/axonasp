<%@ Language=VBScript %>
<%
Response.ContentType = "text/plain"
On Error Resume Next

Sub PrintResult(label)
    Response.Write label & "|Err=" & Err.Number & "|Desc=" & Err.Description & vbCrLf
    Err.Clear
End Sub

Dim obj

Set obj = Server.CreateObject("ADODB.Command")
If IsObject(obj) Then
    obj.CommandText = "SELECT ?"
    obj.Parameters.Append obj.CreateParameter("p1", 200, 1, 10, "ok")
    Response.Write "ADODB.Command.Parameters.Count=" & obj.Parameters.Count & vbCrLf
End If
Call PrintResult("ADODB.Command")

Set obj = Server.CreateObject("Microsoft.XMLDOM")
If IsObject(obj) Then
    obj.LoadXML "<root><item>1</item></root>"
    Response.Write "Microsoft.XMLDOM.DocumentElement=" & obj.DocumentElement.NodeName & vbCrLf
End If
Call PrintResult("Microsoft.XMLDOM")

Set obj = Server.CreateObject("MSXML2.DOMDocument.3.0")
If IsObject(obj) Then
    obj.LoadXML "<root><item>2</item></root>"
    Response.Write "MSXML2.DOMDocument.3.0.DocumentElement=" & obj.DocumentElement.NodeName & vbCrLf
End If
Call PrintResult("MSXML2.DOMDocument.3.0")

Set obj = Server.CreateObject("MSXML2.ServerXMLHTTP")
Call PrintResult("MSXML2.ServerXMLHTTP")

Set obj = Server.CreateObject("MSXML2.XMLHTTP.6.0")
Call PrintResult("MSXML2.XMLHTTP.6.0")

Set obj = Server.CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
If IsObject(obj) Then
    Dim hashBytes
    hashBytes = obj.ComputeHash("axonasp")
    Response.Write "MD5.HashSize=" & obj.HashSize & vbCrLf
End If
Call PrintResult("System.Security.Cryptography.MD5CryptoServiceProvider")
%>
