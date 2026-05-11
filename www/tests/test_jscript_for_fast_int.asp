<script runat="server" language="JScript">
    function runLoop() {
        var sum = 0;
        for (let i = 0; i < 100; i++) {
            sum = sum + i;
        }
        return sum;
    }

    Response.Write(runLoop());
</script>