<% 
' Test for form class method calls with arguments
Set aspForm = New aspForm

Response.Write("<h2>Testing aspForm.build() method</h2>")

' Test 1: Check if id variable is passed correctly to build method
id = "testHashForm"
Response.Write("<p>Test 1: Testing aspForm.build with id='testHashForm'</p>")
result = aspForm.build(id)
Response.Write("<p>Result from build: " & CStr(result) & "</p>")

' Test 2: Check md5 hash functionality
Response.Write("<p>Test 2: Testing hash generation</p>")
testValue = "Hello"
Set objCrypto = Server.CreateObject("G3CRYPTO")
hashResult = objCrypto.HashPassword(testValue)
Response.Write("<p>MD5 result: " & hashResult & "</p>")

Response.Write("<p>Tests completed successfully!</p>")
%>
