<%
@ Language = "JScript"
%>
<script runat="server" language="JScript">
    var s = Symbol("secret");
    var proto = { x: 7 };
    var obj = {};
    Object.setPrototypeOf(obj, proto);
    obj[s] = 42;

    var symbols = Object.getOwnPropertySymbols(obj);
    Response.Write("is=" + (Object.is(NaN, NaN) && !Object.is(0, -0) && Object.is(-0, -0) ? "yes" : "no") + "\n");
    Response.Write("proto=" + (Object.getPrototypeOf(obj) === proto ? "yes" : "no") + "\n");
    Response.Write("symbols=" + symbols.length + "," + (symbols[0] === s ? "yes" : "no") + "\n");
    Response.Write("inherit=" + obj.x);
</script>