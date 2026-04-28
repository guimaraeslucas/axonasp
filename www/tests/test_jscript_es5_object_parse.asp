<%
@ Language = "JScript"
%>
<script runat="server" language="JScript">
    var p1 = parseInt("08");
    var p2 = parseInt("0x10");
    var p3 = parseInt("11", 2);
    var p4 = parseInt("xyz");

    var f1 = parseFloat("3.14abc");
    var f2 = parseFloat(".5");
    var f3 = parseFloat("1e-2");
    var f4 = parseFloat("abc");

    Response.Write(
        "PARSEI=" +
            p1 +
            ":" +
            p2 +
            ":" +
            p3 +
            ":" +
            (p4 != p4 ? "NaN" : p4) +
            ";"
    );
    Response.Write(
        "PARSEF=" +
            f1 +
            ":" +
            f2 +
            ":" +
            f3 +
            ":" +
            (f4 != f4 ? "NaN" : f4) +
            ";"
    );
</script>
