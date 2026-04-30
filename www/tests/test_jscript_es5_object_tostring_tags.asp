<%
@ Language = "JScript"
%>
<script runat="server" language="JScript">
    var out = [];
    out.push({}.toString());
    out.push({ __js_type: "Array" }.toString());
    out.push({ __js_type: "Date" }.toString());
    out.push({ __js_type: "Function" }.toString());
    out.push({ __js_type: "RegExp" }.toString());

    Response.Write(out.join("|"));
</script>
