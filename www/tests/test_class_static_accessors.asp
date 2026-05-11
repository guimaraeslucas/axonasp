<script runat="server" language="JScript">
    class Utils {
        static square(x) {
            return x * x;
        }
        static cube(x) {
            return x * x * x;
        }
    }

    Response.Write("Static methods: ");
    var s = Utils.square(4);
    var c = Utils.cube(3);
    if (s === 16 && c === 27) {
        Response.Write("PASS");
    } else {
        Response.Write("FAIL (expected 16|27, got " + s + "|" + c + ")");
    }
    Response.Write("<br>");

    class Person {
        constructor(name) {
            this._name = name;
        }
        get name() {
            return this._name.toUpperCase();
        }
        set name(value) {
            this._name = value;
        }
    }

    var p = new Person("alice");
    Response.Write("Getter: ");
    if (p.name === "ALICE") {
        Response.Write("PASS");
    } else {
        Response.Write("FAIL (expected ALICE, got " + p.name + ")");
    }
    Response.Write("<br>");

    p.name = "bob";
    Response.Write("Setter: ");
    if (p.name === "BOB") {
        Response.Write("PASS");
    } else {
        Response.Write("FAIL (expected BOB, got " + p.name + ")");
    }
</script>
