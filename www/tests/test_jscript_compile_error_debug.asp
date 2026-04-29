<%
@ Language = "VBScript"
%>
<html>
    <head>
        <title>JScript Compile Error Debug</title>
    </head>
    <body>
        <script language="JScript" runat="server">
            function BrokenCompileBlock() {
                var x = ;
            }
        </script>
        <p>If this page renders normally, the test failed.</p>
    </body>
</html>
