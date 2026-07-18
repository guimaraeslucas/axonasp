# Modern Regular Expressions (RegExp)

AxonASP uses a PCRE-compatible engine for JScript Regular Expressions, supporting advanced features introduced in ES6 and subsequent standards (ES2018+).

## Named Capture Groups

Named capture groups allow you to assign names to capturing groups, which can then be accessed via the `groups` property of the match result.

### Syntax

```javascript
var re = /(?<name>pattern)/;
```

## Lookaround Assertions

Lookaround assertions (lookahead and lookbehind) allow matching a pattern based on what precedes or follows it, without including those characters in the match.

### Syntax

- **Positive Lookahead:** `(?=...)`
- **Negative Lookahead:** `(?!...)`
- **Positive Lookbehind:** `(?<=...)`
- **Negative Lookbehind:** `(?<!...)`

## Sticky Flag (y)

The `y` flag indicates that the match must start exactly at the `lastIndex` property of the regular expression object. If the match fails, `lastIndex` is reset to `0`.

## RegExp.prototype.flags

The `flags` property returns a string containing the flags of the regular expression object, sorted alphabetically (`g`, `i`, `m`, `s`, `u`, `y`).

## Code Example

```javascript
<script runat="server" language="JScript">
// 1. Named Capture Groups
var re = /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/;
var match = re.exec("2026-05-14");
Response.Write(match.groups.year); // Output: 2026

// 2. Lookbehind
var reLookbehind = /(?<=\$)\d+/;
Response.Write(reLookbehind.exec("Price: $100")[0]); // Output: 100

// 3. Sticky Flag
var reSticky = /a/y;
reSticky.lastIndex = 1;
Response.Write(reSticky.exec("ba") !== null); // Output: true

// 4. Flags Property
Response.Write(/abc/gimuy.flags); // Output: gimuy
</script>
```
