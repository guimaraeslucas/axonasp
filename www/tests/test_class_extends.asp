<script runat="server" language="JScript">
    function assert(condition, message) {
        if (!condition) {
            Response.Write("FAIL: " + message + "<br>");
        }
    }

    class Animal {
        constructor(name) {
            this.name = name;
        }
        speak() {
            return this.name + " makes a noise";
        }
    }

    class Dog extends Animal {
        constructor(name) {
            super(name);
        }
        speak() {
            return this.name + " barks";
        }
    }

    var d = new Dog("Mitzie");
    assert(d instanceof Dog, "d should be an instance of Dog");
    assert(d instanceof Animal, "d should inherit from Animal");
    assert(d.speak() === "Mitzie barks", "d.speak() should call the overridden method");
    Response.Write("Inheritance: PASS<br>");

    class Base {
        static staticMethod() {
            return "static";
        }
    }
    class Derived extends Base { }

    assert(Derived.staticMethod() === "static", "Derived should inherit staticMethod()");
    Response.Write("Static inheritance: PASS<br>");

</script>