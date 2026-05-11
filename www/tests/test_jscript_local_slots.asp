<script runat="server" language="JScript">
    function LocalSlotCounter() {
        var i = 0;
        for (i = 0; i < 5; i++) {
        }
        return i;
    }

    function LetShadowingCheck() {
        var x = 10;
        {
            let x = 20;
            x = 30;
        }
        return x;
    }

    Response.Write("local_loop=" + LocalSlotCounter() + ";");
    Response.Write("let_shadow=" + LetShadowingCheck() + ";");
</script>