<%
' Test AxMin and AxMax functions
Response.Write "Testing AxMin and AxMax functions<br>"
Response.Write "AxMax(5, 12, 3, 8) = " & AxMax(5, 12, 3, 8) & "<br>"
Response.Write "AxMin(5, 12, 3, 8) = " & AxMin(5, 12, 3, 8) & "<br>"

' Test AxExplode
Response.Write "AxExplode test<br>"
Dim parts
parts = AxExplode(",", "apple,banana,orange")
Response.Write "Parts count: " & UBound(parts) + 1 & "<br>"

' Test AxMD5
Response.Write "AxMD5('hello') = " & AxMD5("hello") & "<br>"

' Test AxSHA1
Response.Write "AxSHA1('hello') = " & AxSHA1("hello") & "<br>"

' Test AxBase64Encode/Decode
Dim encoded, decoded
encoded = AxBase64Encode("Hello World")
decoded = AxBase64Decode(encoded)
Response.Write "Encoded: " & encoded & "<br>"
Response.Write "Decoded: " & decoded & "<br>"

' Test AxHtmlSpecialChars
Response.Write "AxHtmlSpecialChars('<p>Test</p>') = " & AxHtmlSpecialChars("<p>Test</p>") & "<br>"

' Test AxIsInt
Response.Write "AxIsInt(5) = " & AxIsInt(5) & "<br>"
Response.Write "AxIsInt(5.5) = " & AxIsInt(5.5) & "<br>"

' Test AxEmpty
Response.Write "AxEmpty('') = " & AxEmpty("") & "<br>"
Response.Write "AxEmpty('hello') = " & AxEmpty("hello") & "<br>"

Response.Write "<br>All tests completed!"
%>
