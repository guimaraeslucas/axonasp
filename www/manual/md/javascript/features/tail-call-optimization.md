# Tail Call Optimization (TCO)

The use of Tail Call Optimization **MUST BE AVOIDED** whenever possible. Although the AxonASP implementation includes several mechanisms to prevent memory overflow, the risk generally outweighs the minor gains in performance. A script not well formed can easily consume up to 4GB of memory per execution if implemented incorrectly. Tail Call Optimization is an ECMAScript 6 requirement, but even engines like V8 (Chrome) do not implement it due to the problems TCO can cause. If you choose to implement this type of code, test it thoroughly with the server to avoid the risk of exhausting your service/container memory.

## Syntax

```javascript
function sum(n, acc) {
    if (n === 0) {
        return acc;
    }
    return sum(n - 1, acc + n);
}
```

## Remarks

- Tail-position calls in `return` statements are optimized by the JScript VM to reuse the active function frame.
- The optimization currently applies to direct calls (`return fn(...)`) and member calls (`return obj.fn(...)`).
- Tail-call optimization is intentionally disabled when the `return` statement is inside `try`, `catch`, or `finally` blocks to preserve exception-handler semantics.
- If the tail-position call target resolves to a native host function, the VM executes it as a normal call and returns the result without frame reuse.
- Tail Call Optimization enforces a limit of 10,000 instructions; exceeding this threshold safely and automatically triggers a Stack Overflow error, halting script execution. This mechanism ensures memory stability and prevents malicious scripts from exhausting server memory.

## Code Example

```javascript
<script runat="server" language="JScript">
function sum(n, acc) {
    if (n === 0) {
        return acc;
    }
    return sum(n - 1, acc + 1);
}

Response.Write(sum(100000, 0));
// Output: 100000
</script>
```
