# DateTime Functions & Const Declarations

This document describes the DateTime functions and Const declarations implemented in G3 AxonASP.

## Const Declarations

Const declarations create read-only constant values that cannot be reassigned.

### Syntax
```vbscript
Const CONSTANT_NAME = expression
```

### Examples
```vbscript
' Const declarations
Const MAX_ITEMS = 100
Const APP_NAME = "G3 AxonASP"
Const PI = 3.14159
Const TIMEOUT = 30

' Constants cannot be reassigned (error if attempted)
' CONST MAX_ITEMS = 200  ' This would cause an error
```

### Features
- Constants are **read-only** - reassignment attempts will fail
- Constants follow VBScript naming conventions (case-insensitive)
- Constants can hold any data type (numeric, string, date, etc.)
- Constants are block-scoped (procedure or module level)

---

## Current Date/Time Functions

### NOW()
Returns the current date and time.

**Syntax:**
```vbscript
variant = NOW()
```

**Example:**
```vbscript
Dim currentDateTime
currentDateTime = NOW()
Response.Write currentDateTime  ' e.g., "2024-03-15 14:30:45"
```

### DATE()
Returns the current date (without time).

**Syntax:**
```vbscript
variant = DATE()
```

**Example:**
```vbscript
Dim today
today = DATE()
Response.Write today  ' e.g., "2024-03-15"
```

### TIME()
Returns the current time (without date).

**Syntax:**
```vbscript
variant = TIME()
```

**Example:**
```vbscript
Dim currentTime
currentTime = TIME()
Response.Write currentTime  ' e.g., "14:30:45"
```

---

## Date/Time Parsing Functions

### DateValue(dateString)
Converts a string to a date value. Supports multiple date formats:
- ISO 8601: "2024-03-15", "2024-03-15T14:30:45"
- US Format: "12/25/2024", "12/25/2024 14:30:45"
- Text Format: "March 15, 2024"

**Syntax:**
```vbscript
variant = DateValue(dateString)
```

**Example:**
```vbscript
Dim date1, date2, date3
date1 = DateValue("2024-03-15")
date2 = DateValue("12/25/2024")
date3 = DateValue("March 15, 2024")
Response.Write date1  ' "2024-03-15"
```

### TimeValue(timeString)
Converts a string to a time value. Supports multiple formats:
- 24-hour: "14:30:45", "14:30", "14"
- 12-hour: "2:30:45 PM", "02:45:30 AM"

**Syntax:**
```vbscript
variant = TimeValue(timeString)
```

**Example:**
```vbscript
Dim time1, time2
time1 = TimeValue("14:30:45")
time2 = TimeValue("02:45:30 PM")
Response.Write time1  ' "14:30:45"
```

---

## Date Construction Functions

### DateSerial(year, month, day)
Constructs a date from year, month, and day components.

**Syntax:**
```vbscript
variant = DateSerial(year, month, day)
```

**Parameters:**
- **year**: Integer year (1900-9999)
- **month**: Integer month (1-12)
- **day**: Integer day (1-31)

**Example:**
```vbscript
Dim birthDate
birthDate = DateSerial(1990, 5, 20)
Response.Write birthDate  ' "1990-05-20"
```

### TimeSerial(hour, minute, second)
Constructs a time from hour, minute, and second components.

**Syntax:**
```vbscript
variant = TimeSerial(hour, minute, second)
```

**Parameters:**
- **hour**: Integer hour (0-23)
- **minute**: Integer minute (0-59)
- **second**: Integer second (0-59)

**Example:**
```vbscript
Dim eventTime
eventTime = TimeSerial(14, 30, 45)
Response.Write eventTime  ' "14:30:45"
```

---

## Date/Time Extraction Functions

### Year(date)
Extracts the year from a date value.

**Example:**
```vbscript
Response.Write Year(DateValue("2024-03-15"))  ' 2024
```

### Month(date)
Extracts the month from a date value (1-12).

**Example:**
```vbscript
Response.Write Month(DateValue("2024-03-15"))  ' 3
```

### Day(date)
Extracts the day from a date value (1-31).

**Example:**
```vbscript
Response.Write Day(DateValue("2024-03-15"))  ' 15
```

### Hour(time)
Extracts the hour from a time value (0-23).

**Example:**
```vbscript
Response.Write Hour(TimeValue("14:30:45"))  ' 14
```

### Minute(time)
Extracts the minute from a time value (0-59).

**Example:**
```vbscript
Response.Write Minute(TimeValue("14:30:45"))  ' 30
```

### Second(time)
Extracts the second from a time value (0-59).

**Example:**
```vbscript
Response.Write Second(TimeValue("14:30:45"))  ' 45
```

### Weekday(date, [firstDayOfWeek])
Extracts the day of week from a date (1-7, default Sunday=1).

**Syntax:**
```vbscript
integer = Weekday(date, [firstDayOfWeek])
```

**Parameters:**
- **date**: Date value to extract from
- **firstDayOfWeek**: Optional, day to start week (1=Sunday, 2=Monday, etc.)

**Example:**
```vbscript
Response.Write Weekday(DateValue("2024-03-15"))  ' 5 (Friday)
```

---

## Date Arithmetic Functions

### DateAdd(interval, number, date)
Adds a time interval to a date.

**Syntax:**
```vbscript
variant = DateAdd(interval, number, date)
```

**Interval Values:**
- `"yyyy"` - Year
- `"q"` - Quarter
- `"m"` - Month
- `"y"` - Day of year
- `"d"` - Day
- `"w"` - Weekday
- `"ww"` - Week of year
- `"h"` - Hour
- `"n"` - Minute
- `"s"` - Second

**Example:**
```vbscript
Dim baseDate, futureDate
baseDate = DateValue("2024-03-15")
futureDate = DateAdd("m", 2, baseDate)
Response.Write futureDate  ' "2024-05-15"
```

### DateDiff(interval, date1, date2)
Calculates the difference between two dates.

**Syntax:**
```vbscript
long = DateDiff(interval, date1, date2)
```

**Interval Values:** Same as DateAdd

**Example:**
```vbscript
Dim startDate, endDate, days
startDate = DateValue("2024-01-01")
endDate = DateValue("2024-12-31")
days = DateDiff("d", startDate, endDate)
Response.Write days  ' 364 (days between dates)
```

---

## Date Part Extraction Function

### DatePart(interval, date, [firstDayOfWeek])
Extracts a specific part of a date.

**Syntax:**
```vbscript
integer = DatePart(interval, date, [firstDayOfWeek])
```

**Interval Values:**
- `"yyyy"` - Year
- `"q"` - Quarter (1-4)
- `"m"` - Month (1-12)
- `"y"` - Day of year (1-366)
- `"d"` - Day (1-31)
- `"w"` - Weekday (1-7)
- `"ww"` - Week of year (1-53)
- `"h"` - Hour (0-23)
- `"n"` - Minute (0-59)
- `"s"` - Second (0-59)

**Example:**
```vbscript
Dim date1
date1 = DateValue("2024-03-15")
Response.Write DatePart("q", date1)   ' 1 (Q1)
Response.Write DatePart("ww", date1)  ' 11 (week 11)
Response.Write DatePart("yyyy", date1) ' 2024
```

---

## Date Formatting Function

### FormatDateTime(date, format)
Formats a date/time value for display.

**Syntax:**
```vbscript
string = FormatDateTime(date, format)
```

**Format Values:**
- `0` - **vbGeneralDate** - Default format (e.g., "3/15/2024 2:30:45 PM")
- `1` - **vbLongDate** - Long date format (e.g., "Friday, March 15, 2024")
- `2` - **vbShortDate** - Short date format (e.g., "3/15/2024")
- `3` - **vbLongTime** - Long time format (e.g., "2:30:45 PM")
- `4` - **vbShortTime** - Short time format (e.g., "14:30")

**Example:**
```vbscript
Dim now
now = NOW()
Response.Write FormatDateTime(now, 0)  ' "3/15/2024 2:30:45 PM"
Response.Write FormatDateTime(now, 1)  ' "Friday, March 15, 2024"
Response.Write FormatDateTime(now, 2)  ' "3/15/2024"
Response.Write FormatDateTime(now, 3)  ' "2:30:45 PM"
Response.Write FormatDateTime(now, 4)  ' "14:30"
```

---

## Practical Examples

### Example 1: Age Calculation
```vbscript
Const BIRTHDAY = "1990-05-20"
Dim birthDate, today, age

birthDate = DateValue(BIRTHDAY)
today = DATE()

age = Year(today) - Year(birthDate)
If Month(today) < Month(birthDate) Or _
   (Month(today) = Month(birthDate) And Day(today) < Day(birthDate)) Then
    age = age - 1
End If

Response.Write "Your age: " & age
```

### Example 2: Business Days Calculator
```vbscript
Function BusinessDaysBetween(startDate, endDate)
    Dim days, current
    days = 0
    current = startDate
    
    While current < endDate
        If Weekday(current) > 1 And Weekday(current) < 7 Then
            days = days + 1
        End If
        current = DateAdd("d", 1, current)
    Wend
    
    BusinessDaysBetween = days
End Function
```

### Example 3: Event Reminder
```vbscript
Const EVENT_DATE = "2024-12-25"
Dim eventDate, today, daysUntil

eventDate = DateValue(EVENT_DATE)
today = DATE()
daysUntil = DateDiff("d", today, eventDate)

If daysUntil > 0 Then
    Response.Write "Event in " & daysUntil & " days"
ElseIf daysUntil = 0 Then
    Response.Write "Event is today!"
Else
    Response.Write "Event was " & Abs(daysUntil) & " days ago"
End If
```

### Example 4: Report Header with Date
```vbscript
Dim reportDate
reportDate = NOW()
Response.Write "Report Generated: " & FormatDateTime(reportDate, 1)
Response.Write "Time: " & FormatDateTime(reportDate, 3)
```

---

## VBScript Compatibility

All DateTime functions and Const declarations follow classic VBScript specifications:
- Date values are stored as floating-point numbers (serial dates)
- Time operations are performed with millisecond precision
- DateDiff calculations account for daylight saving time boundaries
- FormatDateTime respects system locale settings for output
- Const declarations enforce strict read-only semantics
- Case-insensitive identifiers per VBScript standards

---

## Test Page

For comprehensive testing of all DateTime functions and Const declarations, see:
`/test_const_datetime.asp`

This page includes:
- Const declaration examples and reassignment protection tests
- All current date/time function examples
- Date/time parsing and construction tests
- Date extraction function demonstrations
- Date arithmetic (DateAdd, DateDiff) examples
- DatePart extraction examples
- FormatDateTime formatting demonstrations
- Practical combined examples (birthday countdown, etc.)
