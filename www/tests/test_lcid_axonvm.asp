<%
' Test LCID propagation in AxonVM pathway
%>
<!DOCTYPE html>
<html>
    <head>
        <title>AxonVM LCID FormatDateTime Test</title>
        <style>
            body {
                font-family: Tahoma;
                margin: 20px;
            }
            .result {
                margin: 10px 0;
                padding: 10px;
                border: 1px solid #ccc;
            }
            .pass {
                background: #e8f5e9;
                color: #2e7d32;
            }
            .code {
                font-family: monospace;
                background: #f5f5f5;
                padding: 2px 5px;
            }
        </style>
    </head>
    <body>
        <h1>AxonVM LCID FormatDateTime Test</h1>

        <div class="result">
            <strong>Test 1: Default LCID (English US - 1033)</strong><br />
            <%
            Response.Write "FormatDateTime(Now(), vbLongTime) with LCID 1033: " & FormatDateTime(Now(), vbLongTime)
            %>
        </div>

        <div class="result">
            <strong>Test 2: Set LCID to Portuguese Brazil (1046)</strong><br />
            <%
            Response.LCID = 1046
            %>
            <%
            Response.Write "FormatDateTime(Now(), vbLongTime) with LCID 1046: " & FormatDateTime(Now(), vbLongTime)
            %>
        </div>

        <div class="result">
            <strong>Test 3: Verify LCID is set</strong><br />
            <%
            Response.Write "Response.LCID = " & Response.LCID
            %>
        </div>

        <div class="result">
            <strong>Test 4: Date format comparison</strong><br />
            <%
            Response.LCID = 1033
            %>
            English (1033) vbLongDate:
            <%
            Response.Write FormatDateTime(Now(), vbLongDate)
            %><br />
            <%
            Response.LCID = 1046
            %>
            Portuguese (1046) vbLongDate:
            <%
            Response.Write FormatDateTime(Now(), vbLongDate)
            %>
        </div>

        <div class="result pass">
            <strong>✓ SUCCESS:</strong> If Portuguese results show 24-hour
            format (HH:MM:SS without AM/PM), and different date format, then
            LCID is properly propagated in AxonVM pathway.
        </div>
    </body>
</html>
