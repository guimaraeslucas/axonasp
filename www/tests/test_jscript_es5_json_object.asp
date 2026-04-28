<%
@ Language = "JScript"
%>
<script runat="server" language="JScript">
    var data = JSON.parse('{"name":"axonasp","v":5}');
    var encoded = JSON.stringify(data);

    var base = { origin: "base" };
    var obj = Object.create(base);
    Object.defineProperty(obj, "x", {
        value: 9,
        writable: false,
        enumerable: true,
        configurable: true,
    });
    Object.defineProperty(obj, "y", {
        value: 4,
        writable: true,
        enumerable: false,
        configurable: true,
    });

    var d = Object.getOwnPropertyDescriptor(obj, "x");
    var keys = Object.keys(obj).join(",");
    var p = Object.getPrototypeOf(obj);

    Response.Write("JSON=" + data.name + ":" + data.v + ";");
    Response.Write("ENC=" + encoded + ";");
    Response.Write(
        "OBJ=" +
            keys +
            ":" +
            d.value +
            ":" +
            (d.writable ? "w" : "nw") +
            ":" +
            p.origin +
            ";"
    );
</script>
