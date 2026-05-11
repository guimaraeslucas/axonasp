<script runat="server" language="JScript">
    class Calculator {
        constructor(x) {
            this.x = x;
        }
        add(y) {
            return this.x + y;
        }
        multiply(y) {
            return this.x * y;
        }
    }

    var calc = new Calculator(10);
    var res1 = calc.add(5);
    var res2 = calc.multiply(3);
    
    Response.Write("Instance methods: ");
    if (res1 === 15 && res2 === 30) {
        Response.Write("PASS");
    } else {
        Response.Write("FAIL (expected 15|30, got " + res1 + "|" + res2 + ")");
    }
    Response.Write("<br>");

    class StrictTest {
        constructor() {
            try {
                undeclaredVar = 1;
                this.strictCtor = false;
            } catch(e) {
                this.strictCtor = true;
            }
        }
        test() {
            try {
                anotherUndeclared = 2;
                return false;
            } catch(e) {
                return true;
            }
        }
    }

    var st = new StrictTest();
    Response.Write("Strict mode in constructor: ");
    Response.Write(st.strictCtor ? "PASS" : "FAIL");
    Response.Write("<br>");
    
    Response.Write("Strict mode in method: ");
    Response.Write(st.test() ? "PASS" : "FAIL");
</script>
