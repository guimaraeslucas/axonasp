<%
@ Language = "JScript"
%>
<script runat="server" language="JScript">
    function Box(v) {
        this.value = v;
    }

    Box.prototype.getValue = function () {
        return this.value;
    };

    Array.prototype.firstItem = function () {
        return this[0];
    };

    var box = new Box(11);
    var values = [42, 99];

    Response.Write("BOX=" + box.getValue() + ";");
    Response.Write("ARR=" + values.firstItem() + ";");
    Response.Write(
        "PROTO=" +
            Object.getPrototypeOf(box).constructor.prototype.getValue.call(
                box
            ) +
            ";"
    );
</script>
