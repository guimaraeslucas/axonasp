# ES6 Classes

AxonASP supports ES6 classes for object-oriented programming. Under the hood, classes are built upon JScript's existing prototype-based inheritance model but with modern syntax and strict semantics.

## Syntax

```javascript
class MyClass [extends BaseClass] {
    constructor(...args) {
        [super(...args);]
        // initialization
    }

    method() { ... }

    static staticMethod() { ... }

    get property() { ... }
    set property(value) { ... }
}
```

## Remarks

- **Strict Mode:** All code within a `class` body (including methods and the constructor) implicitly runs in **Strict Mode**.
- **No Hoisting:** Unlike function declarations, classes are not hoisted. You must declare a class before you can use it (Temporal Dead Zone applies).
- **Instantiation:** Classes must be instantiated with the `new` operator. Calling a class constructor as a normal function (without `new`) throws a `TypeError`.
- **Instance Methods:** Methods defined inside the class are attached to the class's `prototype`.
- **Static Methods:** Methods marked with the `static` keyword are attached directly to the class constructor function.
- **Inheritance:** When a class uses `extends`, AxonASP evaluates the superclass, validates that it is a constructor or `null`, and wires both the constructor chain and the prototype chain.
- **Null Heritage:** `extends null` is supported. In that case, the class prototype chain terminates at `null`.
- **Accessors:** `get` and `set` syntax is supported for defining property getters and setters.
- **Private Fields:** ES2022 private class fields (e.g. `#propertyName`) and private static fields (e.g. `static #staticProperty`) are fully supported. They provide true encapsulation without external memory overhead.

## Code Example

```javascript
<script runat="server" language="JScript">
class Person {
    constructor(name) {
        this._name = name;
    }

    // Instance method
    greet() {
        return "Hello, I'm " + this._name;
    }

    // Static method
    static species() {
        return "Homo Sapiens";
    }

    // Accessors
    get name() {
        return this._name.toUpperCase();
    }

    set name(value) {
        this._name = value;
    }
}

var p = new Person("Alice");
Response.Write(p.greet() + "<br>");       // Output: Hello, I'm Alice
Response.Write(Person.species() + "<br>"); // Output: Homo Sapiens
Response.Write(p.name + "<br>");           // Output: ALICE

p.name = "Bob";
Response.Write(p.name);                   // Output: BOB
</script>
```

## Inheritance with super()

When a class extends another class, you can use the `super()` keyword to invoke the parent class's constructor and `super.method()` to call parent class methods.

### super() in Derived Class Constructors

The `super()` call must be made before accessing `this` in a derived class constructor. If `this` is accessed before `super()` completes, a `ReferenceError` is thrown (Temporal Dead Zone).

```javascript
<script runat="server" language="JScript">
class Animal {
    constructor(name) {
        this.name = name;
    }
    speak() {
        return this.name + " makes a sound";
    }
}

class Dog extends Animal {
    constructor(name, breed) {
        super(name);        // Call parent constructor
        this.breed = breed;
    }
    speak() {
        return super.speak() + " - woof!";
    }
}

var dog = new Dog("Buddy", "Golden Retriever");
Response.Write(dog.speak()); // Output: Buddy makes a sound - woof!
</script>
```

### super.method() Calls

Use `super.method()` to invoke a method from the parent class. This is useful for extending parent behavior without completely overriding it.

```javascript
<script runat="server" language="JScript">
class Calculator {
    add(a, b) {
        return a + b;
    }
}

class AdvancedCalculator extends Calculator {
    add(a, b) {
        var result = super.add(a, b);
        return result + 10; // Add 10 to the base result
    }
}

var calc = new AdvancedCalculator();
Response.Write(calc.add(5, 3)); // Output: 18 (5 + 3 + 10)
</script>
```

### super Property Access

You can also use `super` to set or access properties on the parent class prototype:

```javascript
<script runat="server" language="JScript">
class Base {
    greet() { return "Hello"; }
}

class Derived extends Base {
    greet() {
        return super.greet() + " World";
    }
    setData(val) {
        super.data = val; // Set on instance via parent
    }
}

var d = new Derived();
Response.Write(d.greet());   // Output: Hello World
d.setData(42);
Response.Write(d.data);      // Output: 42
</script>
```

### Remarks for super()

- `super()` **must** be called in a derived class constructor before accessing `this`. Accessing `this` before `super()` throws a `ReferenceError`.
- `super.method()` resolves the method from the parent class's prototype and calls it with the current `this` context.
- Multi-level inheritance is fully supported: `class C extends B extends A` works as expected, with each level able to call its parent via `super`.
- Static methods cannot use `super.method()` unless they are inside a derived static method that explicitly calls a parent static method.
