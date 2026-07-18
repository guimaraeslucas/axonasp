# Iteration Protocol - for...of and Custom Iterables

The iteration protocol allows JScript objects to define or customize their iteration behavior, such as which values are looped over in a `for...of` construct.

## `for...of` Statement

The `for...of` statement creates a loop iterating over iterable objects, including built-in `Array`, `String`, `Set`, `Map`, and custom iterables.

## Built-in Iterables

- **Array**: Iterates over elements.
- **String**: Iterates over characters (handling surrogate pairs).
- **Set**: Iterates over unique values.
- **Map**: Iterates over `[key, value]` entries.

## Custom Iterables

To make an object iterable, it must implement the `[Symbol.iterator]` method, which returns an **Iterator** object. An iterator is an object that has a `next()` method returning an object with two properties: `value` (the next value) and `done` (a boolean indicating completion).

## Code Example

```javascript
<script runat="server" language="JScript">
// 1. Iterate over an Array
var fruits = ["Apple", "Orange", "Banana"];
for (var fruit of fruits) {
    Response.Write(fruit + " "); // Output: Apple Orange Banana 
}

// 2. Manual Iterator usage
var it = fruits[Symbol.iterator]();
var res = it.next();
while (!res.done) {
    Response.Write(res.value + " ");
    res = it.next();
}

// 3. Custom Iterable
var range = {
    from: 1,
    to: 3,
    [Symbol.iterator]: function() {
        return {
            current: this.from,
            last: this.to,
            next: function() {
                if (this.current <= this.last) {
                    return { value: this.current++, done: false };
                } else {
                    return { value: undefined, done: true };
                }
            }
        };
    }
};

for (var n of range) {
    Response.Write(n + " "); // Output: 1 2 3
}
</script>
```
