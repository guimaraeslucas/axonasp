<%
@ Language = "VBScript" CodePage = "65001"
%>
<%
Option Explicit

' ===== Test Response.Cookies.Expires =====

Dim tomorrow, nextWeek, result

' Calculate dates
tomorrow = DateAdd("d", 1, Now())
nextWeek = DateAdd("d", 7, Now())

result = ""

On Error Resume Next

' Test 1: Set cookie with value and expires
Response.Cookies("test_cookie_1") = "testvalue1"
Response.Cookies("test_cookie_1").Expires = tomorrow
If Err.Number = 0 Then
    result = result & "Test 1 - Set Expires (1 day): ✅ OK" & vbCrLf
Else
    result = result & "Test 1 - Set Expires FAILED: " & Err.Description & vbCrLf
    Err.Clear
End If

' Test 2: Set another cookie with different expiry
Response.Cookies("test_cookie_2") = "testvalue2"
Response.Cookies("test_cookie_2").Expires = nextWeek
If Err.Number = 0 Then
    result = result & "Test 2 - Set Expires (7 days): ✅ OK" & vbCrLf
Else
    result = result & "Test 2 - Set Expires FAILED: " & Err.Description & vbCrLf
    Err.Clear
End If

' Test 3: Set cookie domain
Response.Cookies("test_cookie_3") = "testvalue3"
Response.Cookies("test_cookie_3").Domain = "example.com"
If Err.Number = 0 Then
    result = result & "Test 3 - Set Domain: ✅ OK" & vbCrLf
Else
    result = result & "Test 3 - Set Domain FAILED: " & Err.Description & vbCrLf
    Err.Clear
End If

' Test 4: Set cookie path
Response.Cookies("test_cookie_4") = "testvalue4"
Response.Cookies("test_cookie_4").Path = "/app"
If Err.Number = 0 Then
    result = result & "Test 4 - Set Path: ✅ OK" & vbCrLf
Else
    result = result & "Test 4 - Set Path FAILED: " & Err.Description & vbCrLf
    Err.Clear
End If

' Test 5: Set cookie secure flag
Response.Cookies("test_cookie_5") = "testvalue5"
Response.Cookies("test_cookie_5").Secure = True
If Err.Number = 0 Then
    result = result & "Test 5 - Set Secure: ✅ OK" & vbCrLf
Else
    result = result & "Test 5 - Set Secure FAILED: " & Err.Description & vbCrLf
    Err.Clear
End If

' Test 6: Set cookie HttpOnly flag
Response.Cookies("test_cookie_6") = "testvalue6"
Response.Cookies("test_cookie_6").HttpOnly = True
If Err.Number = 0 Then
    result = result & "Test 6 - Set HttpOnly: ✅ OK" & vbCrLf
Else
    result = result & "Test 6 - Set HttpOnly FAILED: " & Err.Description & vbCrLf
    Err.Clear
End If

On Error Goto 0

Response.ContentType = "text/plain"
Response.Charset = "utf-8"
Response.Write result
%>
