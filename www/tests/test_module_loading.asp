<script runat="server" language="JScript">
    import { add, bump, getSeed } from "./test_module_math.js";
    import { render } from "./test_module_bridge.js";

    Response.Write("sum=" + add(2, 3));
    Response.Write("|seed=" + getSeed());
    bump();
    Response.Write("|seed2=" + getSeed());
    Response.Write("|" + render("ok"));
</script>