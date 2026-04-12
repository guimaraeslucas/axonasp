<%@ Language="VBScript" %>
<%
On Error Resume Next
Response.ContentType = "text/html"
Response.Write "<html><body>"
Response.Write "<h1>Class + Dictionary + For Each Test</h1>"

' Simple class to mimic cls_constant
Class TestConstant
    Public iId
    Public sConstant
    Public sValue
    Public bOnline
    Public iType
    Public sParameters
    Public sGlobal
End Class

' Open database
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/db/data_jj2ar6as.mdb")

' Build dictionary with class instances (mimicking constants() function)
Response.Write "<h2>Step 1: Build Dictionary</h2>"
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")

Dim rs
Set rs = conn.Execute("SELECT iId FROM tblConstant WHERE iCustomerID=73 ORDER BY sConstant")
Dim loadCount
loadCount = 0

Do While Not rs.EOF
    Dim tc
    Set tc = New TestConstant
    
    ' Mimic pick() - load constant data
    Dim rsInner
    Set rsInner = conn.Execute("SELECT * FROM tblConstant WHERE iCustomerID=73 AND iId=" & rs("iId"))
    If Not rsInner Is Nothing Then
        If Not rsInner.EOF Then
            tc.iId = rsInner("iId")
            tc.sConstant = rsInner("sConstant")
            tc.sValue = rsInner("sValue")
            tc.bOnline = rsInner("bOnline")
            tc.iType = rsInner("iType")
            tc.sParameters = rsInner("sParameters")
            tc.sGlobal = rsInner("sGlobal")
        End If
        rsInner.Close
        Set rsInner = Nothing
    End If
    
    Response.Write "<div>Loaded: iId=" & tc.iId & " sConstant=" & tc.sConstant & "</div>"
    dict.Add tc.iId, tc
    If Err.Number <> 0 Then
        Response.Write "<div>ERROR adding: " & Err.Description & "</div>"
        Err.Clear
    End If
    
    Set tc = Nothing
    loadCount = loadCount + 1
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

Response.Write "<div><b>Total loaded: " & loadCount & " Dict.Count: " & dict.Count & "</b></div>"

' Step 2: Iterate dictionary (mimicking cacheConstants)
Response.Write "<h2>Step 2: For Each iteration</h2>"
Dim k, iterCount
iterCount = 0
For Each k In dict
    Response.Write "<div>Key=" & k & " dict(k).sConstant=" & dict(k).sConstant & " dict(k).sValue=" & Left(dict(k).sValue & "", 50) & "</div>"
    If Err.Number <> 0 Then
        Response.Write "<div>ERROR accessing: " & Err.Description & "</div>"
        Err.Clear
    End If
    iterCount = iterCount + 1
Next
Response.Write "<div><b>Total iterated: " & iterCount & "</b></div>"

' Step 3: Build 2D array (mimicking cacheConstants exactly)
Response.Write "<h2>Step 3: Build 2D array</h2>"
Dim arrc
ReDim arrc(2, dict.Count)
Dim iRunner
iRunner = 0
For Each k In dict
    arrc(0, iRunner) = dict(k).sConstant
    If dict(k).bOnline Then
        arrc(1, iRunner) = dict(k).sValue
        arrc(2, iRunner) = ""
    Else
        arrc(1, iRunner) = ""
        arrc(2, iRunner) = ""
    End If
    If Err.Number <> 0 Then
        Response.Write "<div>ERROR at iRunner=" & iRunner & ": " & Err.Description & "</div>"
        Err.Clear
    End If
    iRunner = iRunner + 1
Next
Response.Write "<div><b>Array built. iRunner=" & iRunner & "</b></div>"
Response.Write "<div>LBound(arrc,2)=" & LBound(arrc,2) & " UBound(arrc,2)=" & UBound(arrc,2) & "</div>"

' Display array contents
Dim i
For i = 0 To iRunner - 1
    Response.Write "<div>arrc(0," & i & ")=" & arrc(0,i) & " arrc(1," & i & ")=" & Left(arrc(1,i) & "", 50) & "</div>"
Next

conn.Close
Set conn = Nothing

Response.Write "</body></html>"
%>
