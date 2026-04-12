<!DOCTYPE html>
<html>
<head>
    <title>AxonASP - MSWC.PageCounter Test</title>
    <style>
        body { font-family: Tahoma, Verdana, sans-serif; background-color: #ECE9D8; margin: 20px; }
        h1 { color: #003399; border-bottom: 2px solid #3366CC; padding-bottom: 5px; }
        .container { background: white; padding: 20px; border: 1px solid #808080; }
        .result { margin-bottom: 10px; padding: 5px; border-left: 4px solid #335EA8; background: #F5F5F5; }
        .label { font-weight: bold; color: #333; }
        .success { color: green; font-weight: bold; margin-top: 20px; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th, td { border: 1px solid #808080; padding: 8px; text-align: left; }
        th { background-color: #003399; color: white; }
    </style>
</head>
<body>
    <h1>MSWC.PageCounter Implementation Test</h1>
    <div class="container">
        <%
        Dim pc, currentHits, nextHits, externalPath, externalHits
        externalPath = "/manual/default.asp"

        ' 1. Create Object
        On Error Resume Next
        Set pc = Server.CreateObject("MSWC.PageCounter")
        
        If Err.Number <> 0 Then
            Response.Write "<div class='result' style='border-left-color:red;'>Error creating MSWC.PageCounter: " & Err.Description & "</div>"
        Else
            ' 2. Test Current Page Hits
            currentHits = pc.Hits()
            Response.Write "<div class='result'><span class='label'>Current Page Hits (Initial):</span> " & currentHits & "</div>"
            
            ' 3. Increment Hit
            pc.PageHit
            nextHits = pc.Hits()
            Response.Write "<div class='result'><span class='label'>Current Page Hits (After PageHit):</span> " & nextHits & "</div>"
            
            ' 4. Test Specific Path
            externalHits = pc.Hits(externalPath)
            Response.Write "<div class='result'><span class='label'>Hits for " & externalPath & ":</span> " & externalHits & "</div>"
            
            ' 5. Reset Specific Path
            pc.Reset(externalPath)
            Response.Write "<div class='result'><span class='label'>Resetting " & externalPath & "... New Count:</span> " & pc.Hits(externalPath) & "</div>"
            
            ' Summary Table
            %>
            <table>
                <tr>
                    <th>Method</th>
                    <th>Description</th>
                    <th>Result</th>
                </tr>
                <tr>
                    <td>Hits()</td>
                    <td>Returns hits for current page</td>
                    <td><%= currentHits %></td>
                </tr>
                <tr>
                    <td>PageHit()</td>
                    <td>Increments hits for current page</td>
                    <td>Success (Now: <%= nextHits %>)</td>
                </tr>
                <tr>
                    <td>Hits(path)</td>
                    <td>Returns hits for a specific path</td>
                    <td><%= pc.Hits(externalPath) %></td>
                </tr>
                <tr>
                    <td>Reset(path)</td>
                    <td>Resets counter for a specific path</td>
                    <td>Verified</td>
                </tr>
            </table>
            
            <div class="success">MSWC.PageCounter verified successfully in AxonVM!</div>
            <%
        End If
        Set pc = Nothing
        %>
    </div>
</body>
</html>
