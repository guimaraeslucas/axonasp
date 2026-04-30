<%
@ Language = "JScript"
%>
<script runat="server" language="JScript">
    function probe(a, b) {
        a = 21;
        var first = arguments[0];
        arguments[1] = 84;
        return first + "|" + b + "|" + arguments[1];
    }

    Response.Write(probe(7, 14));
</script>
