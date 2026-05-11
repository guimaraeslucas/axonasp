<script runat="server" language="JScript">
    class Base {
        constructor(x) {
            this.x = x;
        }
    }
    class Derived extends Base {
        constructor(x, y) {
            try {
                var test = this.x; // Should throw ReferenceError
                Response.Write("TDZ FAIL");
            } catch(e) {
                Response.Write("TDZ PASS");
            }
            super(x);
            Response.Write("<br>Inside derived constructor, y=" + y);
            this.y = y;
        }
    }

    var d = new Derived(10, 20);
    Response.Write("<br>Super call: ");
    if (d.x === 10 && d.y === 20) {
        Response.Write("PASS");
    } else {
        Response.Write("FAIL (x=" + d.x + ", y=" + d.y + ")");
    }
</script>
