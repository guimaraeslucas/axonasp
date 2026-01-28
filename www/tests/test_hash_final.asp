<%
' Reproduce sampleform7 hash test
dim hash
hash = "password"

dim md5 : set md5 = aspL.plugin("md5")
dim sha256 : set sha256 = aspL.plugin("sha256")

Response.Write "<h3>Hash Test</h3>"
Response.Write "<strong>Input:</strong> " & hash & "<br><br>"

On Error Resume Next
Dim md5Result
md5Result = md5.md5(hash, 32)
If Err.Number <> 0 Then
    Response.Write "<strong>MD5 ERROR:</strong> " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "<strong>MD5 hash:</strong> " & md5Result & "<br>"
    Response.Write "<strong>MD5 Length:</strong> " & Len(md5Result) & "<br>"
End If

Dim sha256Result  
sha256Result = sha256.sha256(hash)
If Err.Number <> 0 Then
    Response.Write "<strong>SHA256 ERROR:</strong> " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "<strong>SHA256 hash:</strong> " & sha256Result & "<br>"
    Response.Write "<strong>SHA256 Length:</strong> " & Len(sha256Result) & "<br>"
End If
On Error Goto 0

Response.Write "<br><em>Expected MD5: 5f4dcc3b5aa765d61d8327deb882cf99</em><br>"
Response.Write "<em>Expected SHA256: 5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8</em>"
%>
