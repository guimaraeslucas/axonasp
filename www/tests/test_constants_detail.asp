<%@ Language="VBScript" %>
<%
Response.ContentType = "text/html"
Response.Write "<html><body>"
Response.Write "<h1>Constants Array Detail Diagnostic</h1>"

' The key is qs_cms_arrconstants73 (confirmed by earlier diagnostic)
Dim arrKey
arrKey = "qs_cms_arrconstants73"
Dim arr
arr = Application(arrKey)

Response.Write "<h2>Step 1: Is the array retrievable?</h2>"
Response.Write "<div>IsEmpty(arr): " & IsEmpty(arr) & "</div>"
Response.Write "<div>IsArray(arr): " & IsArray(arr) & "</div>"
Response.Write "<div>TypeName(arr): " & TypeName(arr) & "</div>"

If Not IsEmpty(arr) Then
    Response.Write "<h2>Step 2: Array Bounds</h2>"
    Dim lb1, ub1, lb2, ub2
    On Error Resume Next
    lb1 = LBound(arr, 1)
    ub1 = UBound(arr, 1)
    lb2 = LBound(arr, 2)
    ub2 = UBound(arr, 2)
    Response.Write "<div>LBound(arr,1) = " & lb1 & "</div>"
    Response.Write "<div>UBound(arr,1) = " & ub1 & "</div>"
    Response.Write "<div>LBound(arr,2) = " & lb2 & "</div>"
    Response.Write "<div>UBound(arr,2) = " & ub2 & "</div>"
    If Err.Number <> 0 Then
        Response.Write "<div>ERROR getting bounds: " & Err.Description & "</div>"
        Err.Clear
    End If
    On Error GoTo 0

    Response.Write "<h2>Step 3: Array Contents (first 30 entries)</h2>"
    Response.Write "<table border='1'><tr><th>Index</th><th>Name (0,i)</th><th>Value (1,i)</th><th>Global (2,i)</th></tr>"
    Dim i, maxShow
    maxShow = ub2
    If maxShow > 29 Then maxShow = 29
    On Error Resume Next
    For i = lb2 To maxShow
        Dim cname, cval, cglob
        cname = arr(0, i)
        cval = arr(1, i)
        cglob = arr(2, i)
        If Err.Number <> 0 Then
            Response.Write "<tr><td>" & i & "</td><td colspan='3'>ERROR: " & Err.Description & "</td></tr>"
            Err.Clear
        Else
            ' Truncate long values for display
            If Len(cval) > 100 Then cval = Left(cval, 100) & "..."
            Response.Write "<tr><td>" & i & "</td><td>" & Server.HTMLEncode(CStr(cname)) & "</td><td>" & Server.HTMLEncode(CStr(cval)) & "</td><td>" & Server.HTMLEncode(CStr(cglob)) & "</td></tr>"
        End If
    Next
    On Error GoTo 0
    Response.Write "</table>"

    Response.Write "<h2>Step 4: Regex Pattern Test on [MENU]</h2>"
    Dim rTest
    Set rTest = New RegExp
    rTest.Global = True
    rTest.IgnoreCase = True
    rTest.Pattern = "\[+(MENU)+[\]]"
    Response.Write "<div>Pattern: " & rTest.Pattern & "</div>"
    Response.Write "<div>Test('[MENU]'): " & rTest.Test("[MENU]") & "</div>"
    Response.Write "<div>Test('hello [MENU] world'): " & rTest.Test("hello [MENU] world") & "</div>"
    Set rTest = Nothing

    Response.Write "<h2>Step 5: Replace Test</h2>"
    Dim testStr
    testStr = "Before [MENU] After"
    Response.Write "<div>Original: " & Server.HTMLEncode(testStr) & "</div>"
    testStr = Replace(testStr, "[MENU]", "REPLACED_MENU", 1, -1, 1)
    Response.Write "<div>After Replace: " & Server.HTMLEncode(testStr) & "</div>"

Else
    Response.Write "<div>Array is empty! Cannot proceed.</div>"
    
    ' Try enumerating all application keys looking for arrconstants
    Response.Write "<h2>Looking for any arrconstants key</h2>"
    Dim allkeys, k
    For Each k In Application.Contents
        If InStr(1, k, "arrconstants", 1) > 0 Then
            Response.Write "<div>Found key: " & k & " = "
            If IsArray(Application.Contents(k)) Then
                Response.Write "ARRAY"
            ElseIf IsObject(Application.Contents(k)) Then
                Response.Write "OBJECT"
            Else
                Response.Write Application.Contents(k)
            End If
            Response.Write "</div>"
        End If
    Next
End If

Response.Write "</body></html>"
%>
