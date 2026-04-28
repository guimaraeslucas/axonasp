<%
@ Language = "JScript"
%>
<script runat="server" language="JScript">
    var obj = {
        count: 1,
        items: [3, 8],
    };

    var first = obj.count++;
    var second = --obj.items[1];
    obj.items[0] = 5;

    Response.Write(
        "UPD=" +
            first +
            ":" +
            obj.count +
            ":" +
            second +
            ":" +
            obj.items[1] +
            ";"
    );
    Response.Write("IDX=" + obj.items[0] + ":" + obj["items"][1] + ";");
</script>
