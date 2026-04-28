<%
@ Language = "JScript"
%>
<script runat="server" language="JScript">
    var a = ["a", "b", "a", "c"];
    var first = a.indexOf("a");
    var last = a.lastIndexOf("a");
    var kind = Array.isArray(a) ? "array" : "other";
    var text = "  AxonASP ES5  ".trim();

    var acc = {
        _v: 0,
        get value() {
            return this._v + 1;
        },
        set value(v) {
            this._v = v;
        },
    };
    acc.value = 41;

    Response.Write("ARR=" + first + ":" + last + ":" + kind + ";");
    Response.Write("TXT=" + text + ";");
    Response.Write("ACC=" + acc.value + ":" + acc._v + ";");
</script>
