# Internationalization API (Intl)

## Syntax

```javascript
var dtf = new Intl.DateTimeFormat(locales[, options]);
var nfmt = new Intl.NumberFormat(locales[, options]);
var coll = new Intl.Collator(locales[, options]);
var plur = new Intl.PluralRules(locales[, options]);
var rtf = new Intl.RelativeTimeFormat(locales[, options]);
```

## Remarks

- `Intl` is available as a global namespace in JScript.
- `DateTimeFormat`, `NumberFormat`, `Collator`, `PluralRules`, and `RelativeTimeFormat` use AxonASP locale profiles and the current server locale when no locale is supplied.
- Locale input can be a string or an array-like value. AxonASP uses the first usable locale tag and falls back to the effective server locale, then `en-US`.
- `Intl.DateTimeFormat` supports `dateStyle`, `timeStyle`, `year`, `month`, `day`, `weekday`, `hour`, `minute`, `second`, `hour12`, and `formatToParts()`.
- `Intl.NumberFormat` supports `style: "decimal"`, `style: "currency"`, `style: "percent"`, and `formatToParts()`.
- `Intl.Collator` supports `usage` ("sort", "search"), `sensitivity` ("base", "accent", "case", "variant"), `numeric`, `caseFirst`, and `ignorePunctuation`.
- `Intl.PluralRules` supports `type` ("cardinal", "ordinal") and provides `select(number)`.
- `Intl.RelativeTimeFormat` supports `numeric` ("always", "auto"), `style` ("long", "short", "narrow"), and provides `format(value, unit)` and `formatToParts(value, unit)`.
- Unsupported locale values and extra options are ignored or fall back to the closest supported locale profile.

## Code Example

```javascript
<script runat="server" language="JScript">
var dateValue = new Date(Date.UTC(2026, 0, 2, 3, 4, 5));
var enDate = new Intl.DateTimeFormat("en-US", { dateStyle: "short" }).format(dateValue);
var ptDate = new Intl.DateTimeFormat("pt-BR", { dateStyle: "short" }).format(dateValue);
var deNumber = new Intl.NumberFormat("de-DE", { style: "currency", currency: "EUR", maximumFractionDigits: 2 }).format(1234567.89);

Response.Write(enDate + "\n");
Response.Write(ptDate + "\n");
Response.Write(deNumber + "\n");
// Output:
// 1/2/2026
// 02/01/2026
// EUR 1.234.567,89

// Collator Example
var collator = new Intl.Collator("en", { sensitivity: "base" });
Response.Write(collator.compare("a", "A") === 0); // Output: true

// PluralRules Example
var pr = new Intl.PluralRules("en");
Response.Write(pr.select(1)); // one
Response.Write(pr.select(2)); // other

// RelativeTimeFormat Example
var rtf = new Intl.RelativeTimeFormat("en", { numeric: "auto" });
Response.Write(rtf.format(-1, "day")); // yesterday
Response.Write(rtf.format(2, "day"));  // in 2 days
</script>
```
