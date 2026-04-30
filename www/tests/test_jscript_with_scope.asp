<%
@ Language = "JScript"
%>
<script runat="server" language="JScript">
    var outside = "global";
    var bag = { outside: "local", count: 1 };
    with (bag) {
        outside = outside + "-" + count;
        count = 9;
    }
    Response.Write(outside + "|" + bag.outside + "|" + bag.count);
</script>
