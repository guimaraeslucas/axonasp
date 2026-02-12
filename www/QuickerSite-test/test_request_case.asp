<%
Response.Write "Test Request Case Sensitivity" & vbCrLf
Response.Write "==============================" & vbCrLf & vbCrLf

' Test lowercase parameter name in code accessing uppercase in URL
Response.Write "Request(""iId""): [" & Request("iId") & "]" & vbCrLf
Response.Write "Request(""iID""): [" & Request("iID") & "]" & vbCrLf
Response.Write "Request(""IID""): [" & Request("IID") & "]" & vbCrLf
Response.Write "Request(""Iid""): [" & Request("Iid") & "]" & vbCrLf

Response.Write vbCrLf
Response.Write "QueryString(""iId""): [" & Request.QueryString("iId") & "]" & vbCrLf
Response.Write "QueryString(""iID""): [" & Request.QueryString("iID") & "]" & vbCrLf

Response.Write vbCrLf
Response.Write "IsEmpty(Request(""iId"")): " & IsEmpty(Request("iId")) & vbCrLf
Response.Write "IsEmpty(Request(""iID"")): " & IsEmpty(Request("iID")) & vbCrLf
Response.Write "Len(Request(""iId"")): " & Len(Request("iId")) & vbCrLf
Response.Write "Len(Request(""iID"")): " & Len(Request("iID")) & vbCrLf
%>
