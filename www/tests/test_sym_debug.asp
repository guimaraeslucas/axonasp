<%@ Language="JScript" %>
<script runat="server" language="JScript">
var sym1 = Symbol("a");
var sym2 = Symbol("b");
var obj = { [sym1]: 1, [sym2]: 2, c: 3 };
var syms = Object.getOwnPropertySymbols(obj);
Response.Write("length=" + syms.length + "\n");
Response.Write("type=" + typeof(syms) + "\n");
if (syms.length > 0) {
    Response.Write("0=" + typeof(syms[0]) + "\n");
}
</script>