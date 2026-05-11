<script runat="server" language="JScript">
    class Base {
        constructor() { this.x = 1; }
    }
    class Derived extends Base {
        constructor() {
            super();
            Response.Write("x=" + this.x);
        }
    }
    new Derived();
</script>
