<%
@ Language = "JScript"
%>
<script runat="server" language="JScript">
    var payload = {
        value: 7,
        toJSON: function () {
            return {
                encoded: this.value * 3,
                label: "custom",
            };
        },
    };

    Response.Write(JSON.stringify(payload));
</script>
