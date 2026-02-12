<%
Dim arr, k, found
found = 0
If IsEmpty(Application("QS_CMS_arrconstants")) Then
  Response.Write "arrconstants is EMPTY"
  Response.End
End If

arr = Application("QS_CMS_arrconstants")
For k = LBound(arr,2) To UBound(arr,2)
  If UCase(CStr(arr(0,k))) = "SAMPLESCRIPT1" Or UCase(CStr(arr(0,k))) = "SAMPLESCRIPT3" Or UCase(CStr(arr(0,k))) = "SAMPLESCRIPT4" Then
    found = found + 1
    Response.Write "---<br/>"
    Response.Write "name=" & arr(0,k) & "<br/>"
    Response.Write "len(valueField)=" & Len(CStr(arr(1,k))) & "<br/>"
    Response.Write "hasIdentifier=" & (InStr(1,CStr(arr(1,k)),"#######",1) > 0) & "<br/>"
    Response.Write "globalLen=" & Len(CStr(arr(2,k))) & "<br/>"
    Response.Write "preview=" & Server.HTMLEncode(Left(CStr(arr(1,k)),200)) & "<br/>"
  End If
Next
Response.Write "FOUND=" & found
%>