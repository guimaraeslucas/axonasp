# Locale Support and Language Formatting

## Overview

AxonASP provides comprehensive locale support for date, time, number, and currency formatting. The server supports 43 language and region variants, enabling global ASP applications to render content in multiple languages and formats without manual conversion.

Locale settings control how the following functions format and parse data:

- **Date/Time Functions:** `FormatDateTime`, `DateValue`, `TimeValue`, `CDate`, `DateAdd`, `DateDiff`, `DatePart`, `Weekday`
- **Number Functions:** `FormatNumber`, `FormatPercent`, `FormatCurrency`
- **Day/Month Names:** `MonthName`, `WeekdayName`
- **Implicit Conversions:** Date and number values converted to strings

## Working with Locales

### Setting the Locale

The locale is controlled by the **Session.LCID** property. By default, it inherits from the server's configured default locale (`global.default_mslcid` in `axonasp.toml`).

**In ASP Code:**

```asp
' Set locale for Portuguese Brazil
Session.LCID = 1046

' Format a number using Portuguese locale
amount = 1234.56
formatted = FormatNumber(amount, 2)  ' Outputs: 1.234,56

' Format currency
currency = FormatCurrency(amount)    ' Outputs: R$ 1.234,56
```

**Checking Current Locale:**

```asp
currentLocale = Session.LCID
Response.Write "Current locale: " & currentLocale
```

### Server Default Locale

The server default is configured in `axonasp.toml`:

```toml
[global]
default_mslcid = 1046  # Portuguese (Brazil)
```

This default applies to all new sessions. Individual requests can override it using `Session.LCID`.

## Supported Locales and Variants

AxonASP supports the following 43 language and region locales:

### English Variants

| Locale | LCID | Region | Decimal | Thousands | Currency |
|--------|------|--------|---------|-----------|----------|
| en-US | 1033 | United States | . | , | $ |
| en-GB | 2057 | United Kingdom | . | , | £ |
| en-AU | 3081 | Australia | . | , | $ |
| en-CA | 4105 | Canada | . | , | $ |
| en-IN | 16393 | India | . | , | ₹ |
| en-IE | 6153 | Ireland | . | , | € |
| en-NZ | 5129 | New Zealand | . | , | $ |
| en-ZA | 7177 | South Africa | . | , | R |

### Portuguese Variants

| Locale | LCID | Region | Decimal | Thousands | Currency |
|--------|------|--------|---------|-----------|----------|
| pt-BR | 1046 | Brazil | , | . | R$ |
| pt-PT | 2070 | Portugal | , | . | € |

### Spanish Variants

| Locale | LCID | Region | Decimal | Thousands | Currency |
|--------|------|--------|---------|-----------|----------|
| es-ES | 1034 | Spain | , | . | € |
| es-MX | 2058 | Mexico | . | , | $ |
| es-AR | 11274 | Argentina | , | . | $ |
| es-CO | 9226 | Colombia | , | . | $ |
| es-CL | 13322 | Chile | , | . | $ |
| es-PE | 10250 | Peru | . | , | S/ |

### French Variants

| Locale | LCID | Region | Decimal | Thousands | Currency |
|--------|------|--------|---------|-----------|----------|
| fr-FR | 1036 | France | , | (space) | € |
| fr-CA | 3084 | Canada | , | (space) | $ |
| fr-BE | 2060 | Belgium | , | (space) | € |
| fr-CH | 4108 | Switzerland | . | ' | CHF |

### German Variants

| Locale | LCID | Region | Decimal | Thousands | Currency |
|--------|------|--------|---------|-----------|----------|
| de-DE | 1031 | Germany | , | . | € |
| de-AT | 3079 | Austria | , | . | € |
| de-CH | 2055 | Switzerland | . | ' | CHF |

### Italian Variants

| Locale | LCID | Region | Decimal | Thousands | Currency |
|--------|------|--------|---------|-----------|----------|
| it-IT | 1040 | Italy | , | . | € |
| it-CH | 4108 | Switzerland | . | ' | CHF |

### Dutch Variants

| Locale | LCID | Region | Decimal | Thousands | Currency |
|--------|------|--------|---------|-----------|----------|
| nl-NL | 1043 | Netherlands | , | . | € |
| nl-BE | 2060 | Belgium | , | . | € |

### Scandinavian Variants

| Locale | LCID | Region | Decimal | Thousands | Currency |
|--------|------|--------|---------|-----------|----------|
| da-DK | 1030 | Denmark | , | . | kr |
| fi-FI | 1035 | Finland | , | (space) | € |
| nb-NO | 1044 | Norway | , | (space) | kr |

### Slavic Variants

| Locale | LCID | Region | Decimal | Thousands | Currency |
|--------|------|--------|---------|-----------|----------|
| pl-PL | 1045 | Poland | , | (space) | zł |
| cs-CZ | 1029 | Czech Republic | , | (space) | Kč |
| bg-BG | 1026 | Bulgaria | , | (space) | лв |
| ru-RU | 1049 | Russia | , | (space) | ₽ |
| uk-UA | 1058 | Ukraine | , | (space) | ₴ |
| sk-SK | 1051 | Slovakia | , | (space) | € |
| hr-HR | 1050 | Croatia | , | . | € |

### Asian Variants

| Locale | LCID | Region | Decimal | Thousands | Currency |
|--------|------|--------|---------|-----------|----------|
| zh-CN | 2052 | China | . | , | ¥ |
| zh-TW | 1028 | Taiwan | . | , | NT$ |
| zh-HK | 3076 | Hong Kong | . | , | HK$ |
| ja-JP | 1041 | Japan | . | , | ¥ |
| ko-KR | 1042 | Korea | . | , | ₩ |
| th-TH | 1054 | Thailand | . | , | ฿ |

### Other Variants

| Locale | LCID | Region | Decimal | Thousands | Currency |
|--------|------|--------|---------|-----------|----------|
| el-GR | 1032 | Greece | , | . | € |
| id-ID | 1057 | Indonesia | , | . | Rp |
| tr-TR | 1055 | Turkey | , | . | ₺ |

## Formatting Examples

### Number Formatting

Numbers are formatted with locale-specific decimal and thousands separators:

**ASP Code:**

```asp
Dim amount, formatted

amount = 1234.567
Session.LCID = 1033  ' English US

formatted = FormatNumber(amount, 2)     ' Output: 1,234.57
formatted = FormatPercent(0.125, 1)     ' Output: 12.5%

' Switch to German
Session.LCID = 1031
formatted = FormatNumber(amount, 2)     ' Output: 1.234,57
formatted = FormatPercent(0.125, 1)     ' Output: 12,5%

' Switch to French
Session.LCID = 1036
formatted = FormatNumber(amount, 2)     ' Output: 1 234,57
formatted = FormatPercent(0.125, 1)     ' Output: 12,5%
```

### Currency Formatting

Currency symbols and separators are locale-aware:

```asp
Dim price

price = 1234.50
Session.LCID = 1033  ' English US
Response.Write FormatCurrency(price)    ' Output: $1,234.50

Session.LCID = 1046  ' Portuguese Brazil
Response.Write FormatCurrency(price)    ' Output: R$ 1.234,50

Session.LCID = 1036  ' French France
Response.Write FormatCurrency(price)    ' Output: € 1 234,50
```

### Date Formatting

Dates are formatted according to locale conventions:

```asp
Dim eventDate

eventDate = CDate("2026-04-09")
Session.LCID = 1033  ' English US
Response.Write FormatDateTime(eventDate, 1)  ' Output: Thursday, April 9, 2026

Session.LCID = 1046  ' Portuguese Brazil
Response.Write FormatDateTime(eventDate, 1)  ' Output: quinta-feira, 9 de abril de 2026

Session.LCID = 1036  ' French France
Response.Write FormatDateTime(eventDate, 1)  ' Output: jeudi 9 avril 2026
```

### Month and Day Names

Month and weekday names adapt to the session locale:

```asp
Session.LCID = 1046  ' Portuguese Brazil

Response.Write MonthName(4)         ' Output: abril
Response.Write MonthName(4, True)   ' Output: abr
Response.Write WeekdayName(1)       ' Output: domingo
Response.Write WeekdayName(1, True) ' Output: dom
```

### Date Parsing

Date parsing respects locale conventions when parsing string dates:

```asp
Dim dateStr

Session.LCID = 1046  ' Portuguese Brazil
dateStr = "09/04/2026"  ' Day/Month/Year format
Response.Write CDate(dateStr)  ' Parses as: April 9, 2026

Session.LCID = 1033  ' English US
dateStr = "04/09/2026"  ' Month/Day/Year format
Response.Write CDate(dateStr)  ' Parses as: April 9, 2026
```

## Affected Functions

The following VBScript functions respect the current `Session.LCID`:

### Date/Time Functions

- `FormatDateTime(expression, format)` - All format types (1-4)
- `DateValue(expression)` - Parses dates in locale format
- `TimeValue(expression)` - Parses times
- `CDate(expression)` - Parses dates and times
- `DateAdd(interval, number, date)` - Returns date/time values
- `DateDiff(interval, date1, date2)` - Returns intervals
- `DatePart(interval, date)` - Extracts date parts
- `Day(date)`, `Month(date)`, `Year(date)` - Date component extraction
- `Weekday(date)` - Returns day of week (1 = Sunday, 7 = Saturday)
- `MonthName(number, abbrev)` - Locale-aware month names
- `WeekdayName(index, abbrev)` - Locale-aware weekday names

### Number Functions

- `FormatNumber(expression, numDigits, useGrouping)` - Uses locale separators
- `FormatPercent(expression, numDigits, useGrouping)` - Uses locale separators
- `FormatCurrency(expression, numDigits, useGrouping)` - Uses locale symbol and separators

### String Conversion

- Implicit VTDate to string conversion respects locale date/time format
- Implicit numeric to string conversion uses locale separators

## Performance Notes

Locale resolution is optimized for performance:

- Locale profiles are precomputed at server startup
- Language matching uses efficient BCP 47 tag matching
- Formatting functions cache locale profiles in the VM context
- No reflection or dynamic dispatch is used in formatting paths

## Configuration

The server default locale is configured in `axonasp.toml`:

```toml
[global]
default_mslcid = 1046        # Portuguese (Brazil)
default_timezone = "America/Sao_Paulo"
```

The `default_timezone` setting affects date/time implicit conversions and parsing. Both UTC and named timezones (IANA timezone database) are supported.

## Troubleshooting

### Incorrect Date Parsing

If dates parse incorrectly, verify the session locale matches the date format:

```asp
' Bad: locale mismatch
Session.LCID = 1033  ' US (MM/DD/YYYY)
Response.Write CDate("25/12/2026")  ' Fails - US format expects MM/DD

' Good: locale matches format
Session.LCID = 1046  ' Brazil (DD/MM/YYYY)
Response.Write CDate("25/12/2026")  ' Works correctly
```

### Missing Locale

If a specific locale code is not in the supported list, the system falls back to the closest matching locale (e.g., an unmapped Spanish variant falls back to es-ES).

### Currency Symbol Not Appearing

Ensure the locale is set before calling `FormatCurrency`:

```asp
Session.LCID = 1046  ' Must be set before formatting
Response.Write FormatCurrency(100)
```
