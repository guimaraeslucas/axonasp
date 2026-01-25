/*
 * AxonASP Server - Version 1.0
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * Attribution Notice:
 * If this software is used in other projects, the name "AxonASP Server"
 * must be cited in the documentation or "About" section.
 *
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 */
package server

import (
	"fmt"
	"strings"
	"time"
)

// ConstValue represents a constant value wrapper to prevent modification
type ConstValue struct {
	Value interface{}
	Name  string
}

// DateTimeFormat constants for FormatDateTime
const (
	FormatDateTimeGeneralDate     = 0 // m/d/yy h:mm:ss
	FormatDateTimeLongDate        = 1 // dddd, mmmm dd, yyyy
	FormatDateTimeShortDate       = 2 // m/d/yy
	FormatDateTimeLongTime        = 3 // h:mm:ss AM/PM
	FormatDateTimeShortTime       = 4 // h:mm AM/PM
	FormatDateTimeISOWeekNumber   = 5 // Week number
	FormatDateTimeFirstWeekOfYear = 6 // First week of year
)

// DatePartInterval represents the interval for DatePart/DateAdd
type DatePartInterval string

const (
	DatePartYear       DatePartInterval = "yyyy"
	DatePartQuarter    DatePartInterval = "q"
	DatePartMonth      DatePartInterval = "m"
	DatePartDayOfYear  DatePartInterval = "y"
	DatePartDay        DatePartInterval = "d"
	DatePartWeekday    DatePartInterval = "w"
	DatePartWeekOfYear DatePartInterval = "ww"
	DatePartHour       DatePartInterval = "h"
	DatePartMinute     DatePartInterval = "n"
	DatePartSecond     DatePartInterval = "s"
)

// evalDateTimeFunction evaluates VBScript date/time functions
// Returns (result, wasHandled) where wasHandled indicates if the function was recognized
func evalDateTimeFunction(funcName string, args []interface{}, ctx *ExecutionContext) (interface{}, bool) {
	funcLower := strings.ToLower(funcName)

	switch funcLower {
	case "now":
		// NOW() - returns current date and time
		return time.Now(), true

	case "date":
		// DATE() - returns current date (no time)
		now := time.Now()
		return time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, now.Location()), true

	case "time":
		// TIME() - returns current time (midnight + time)
		now := time.Now()
		return time.Date(1899, time.December, 30, now.Hour(), now.Minute(), now.Second(), 0, now.Location()), true

	case "datevalue":
		// DATEVALUE(string) - converts string to date
		if len(args) == 0 {
			return time.Time{}, true
		}
		dateStr := toString(args[0])
		return parseVBScriptDate(dateStr), true

	case "timevalue":
		// TIMEVALUE(string) - converts string to time
		if len(args) == 0 {
			return time.Time{}, true
		}
		timeStr := toString(args[0])
		return parseVBScriptTime(timeStr), true

	case "dateserial":
		// DATESERIAL(year, month, day) - creates date from components
		if len(args) < 3 {
			return time.Time{}, true
		}
		year := toInt(args[0])
		month := toInt(args[1])
		day := toInt(args[2])
		return time.Date(year, time.Month(month), day, 0, 0, 0, 0, time.UTC), true

	case "timeserial":
		// TIMESERIAL(hour, minute, second) - creates time from components
		if len(args) < 3 {
			return time.Time{}, true
		}
		hour := toInt(args[0])
		minute := toInt(args[1])
		second := toInt(args[2])
		// VBScript represents time as days since Dec 30, 1899 (Ole Automation Date)
		return time.Date(1899, time.December, 30, hour, minute, second, 0, time.UTC), true

	case "year":
		// YEAR(date) - extracts year
		if len(args) == 0 {
			dt := time.Now()
			return dt.Year(), true
		}
		dt := toDateTime(args[0])
		return dt.Year(), true

	case "month":
		// MONTH(date) - extracts month (1-12)
		if len(args) == 0 {
			dt := time.Now()
			return int(dt.Month()), true
		}
		dt := toDateTime(args[0])
		return int(dt.Month()), true

	case "day":
		// DAY(date) - extracts day of month
		if len(args) == 0 {
			dt := time.Now()
			return dt.Day(), true
		}
		dt := toDateTime(args[0])
		return dt.Day(), true

	case "hour":
		// HOUR(time) - extracts hour (0-23)
		if len(args) == 0 {
			dt := time.Now()
			return dt.Hour(), true
		}
		dt := toDateTime(args[0])
		return dt.Hour(), true

	case "minute":
		// MINUTE(time) - extracts minute (0-59)
		if len(args) == 0 {
			dt := time.Now()
			return dt.Minute(), true
		}
		dt := toDateTime(args[0])
		return dt.Minute(), true

	case "second":
		// SECOND(time) - extracts second (0-59)
		if len(args) == 0 {
			dt := time.Now()
			return dt.Second(), true
		}
		dt := toDateTime(args[0])
		return dt.Second(), true

	case "weekday":
		// WEEKDAY(date, [firstDayOfWeek]) - returns day of week (1=Sunday by default)
		if len(args) == 0 {
			return int(time.Now().Weekday()) + 1, true
		}
		dt := toDateTime(args[0])
		firstDay := 1 // Default: Sunday = 1
		if len(args) > 1 {
			firstDay = toInt(args[1])
		}
		// Adjust to VBScript weekday (1=Sunday)
		goWeekday := int(dt.Weekday())
		if firstDay == 1 {
			return goWeekday + 1, true // 0-6 -> 1-7
		}
		// For other firstDayOfWeek values, adjust accordingly
		return ((goWeekday - firstDay + 1) % 7) + 1, true

	case "dayofweek":
		// Alias for WEEKDAY
		if len(args) == 0 {
			return int(time.Now().Weekday()) + 1, true
		}
		dt := toDateTime(args[0])
		return int(dt.Weekday()) + 1, true

	case "dateadd":
		// DATEADD(interval, number, date [, firstdayofweek, firstweekofyear])
		// VBScript ignores the optional params for date math; we accept them for parity.
		if len(args) < 3 {
			return time.Time{}, true
		}
		interval := strings.ToLower(toString(args[0]))
		num := toInt(args[1])
		dt := toDateTime(args[2])

		return addDatePart(dt, interval, num), true

	case "datediff":
		// DATEDIFF(interval, date1, date2 [, firstdayofweek, firstweekofyear])
		if len(args) < 3 {
			return 0, true
		}
		interval := strings.ToLower(toString(args[0]))
		date1 := toDateTime(args[1])
		date2 := toDateTime(args[2])
		firstDay := 1
		firstWeek := 1
		if len(args) >= 4 {
			firstDay = toInt(args[3])
		}
		if len(args) >= 5 {
			firstWeek = toInt(args[4])
		}

		return calcDateDiff(interval, date1, date2, firstDay, firstWeek), true

	case "datepart":
		// DATEPART(interval, date [, firstdayofweek, firstweekofyear])
		if len(args) < 2 {
			return 0, true
		}
		interval := strings.ToLower(toString(args[0]))
		dt := toDateTime(args[1])
		firstDay := 1
		firstWeek := 1
		if len(args) >= 3 {
			firstDay = toInt(args[2])
		}
		if len(args) >= 4 {
			firstWeek = toInt(args[3])
		}

		return extractDatePart(interval, dt, firstDay, firstWeek), true

	case "datefromparts":
		// Non-standard: creates date from parts
		if len(args) < 3 {
			return time.Time{}, true
		}
		year := toInt(args[0])
		month := toInt(args[1])
		day := toInt(args[2])
		return time.Date(year, time.Month(month), day, 0, 0, 0, 0, time.UTC), true

	case "formatdatetime":
		// FORMATDATETIME(date, format) - formats date/time
		if len(args) < 1 {
			return "", true
		}
		dt := toDateTime(args[0])
		format := FormatDateTimeShortDate
		if len(args) > 1 {
			format = toInt(args[1])
		}

		return formatDateTime(dt, format), true

	default:
		return nil, false
	}
}

// Helper functions

// toDateTime converts a value to time.Time
func toDateTime(val interface{}) time.Time {
	if val == nil {
		return time.Now()
	}

	switch v := val.(type) {
	case time.Time:
		return v
	case string:
		// Try parsing
		parsed := parseVBScriptDate(v)
		if !parsed.IsZero() {
			return parsed
		}
		// Try parsing as time
		parsed = parseVBScriptTime(v)
		if !parsed.IsZero() {
			return parsed
		}
		return time.Now()
	case int:
		// VBScript stores dates as days since Dec 30, 1899
		baseDate := time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)
		return baseDate.AddDate(0, 0, v)
	case float64:
		// Float includes time fraction
		baseDate := time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)
		days := int(v)
		fraction := v - float64(days)
		seconds := int(fraction * 86400) // 24*60*60
		return baseDate.AddDate(0, 0, days).Add(time.Duration(seconds) * time.Second)
	default:
		return time.Now()
	}
}

// parseVBScriptDate parses a date string in VBScript format
func parseVBScriptDate(dateStr string) time.Time {
	dateStr = strings.TrimSpace(dateStr)

	formats := []string{
		"1/2/2006",
		"01/02/2006",
		"2006-01-02",
		"1/2/06",
		"01/02/06",
		"2006-1-2",
		"01/02/2006 15:04:05",
		"1/2/2006 3:04:05 PM",
		"January 2, 2006",
		"Jan 2, 2006",
		"2006-01-02 15:04:05",
	}

	for _, format := range formats {
		if t, err := time.Parse(format, dateStr); err == nil {
			return t
		}
	}

	return time.Time{}
}

// parseVBScriptTime parses a time string in VBScript format
func parseVBScriptTime(timeStr string) time.Time {
	timeStr = strings.TrimSpace(timeStr)

	formats := []string{
		"15:04:05",
		"3:04:05 PM",
		"3:04:05 AM",
		"15:04",
		"3:04 PM",
		"3:04 AM",
	}

	for _, format := range formats {
		if t, err := time.Parse(format, timeStr); err == nil {
			return t
		}
	}

	return time.Time{}
}

// addDatePart adds a date part to a time
func addDatePart(dt time.Time, interval string, num int) time.Time {
	switch interval {
	case "yyyy":
		return dt.AddDate(num, 0, 0)
	case "q":
		return dt.AddDate(0, num*3, 0)
	case "m":
		return dt.AddDate(0, num, 0)
	case "y", "d":
		return dt.AddDate(0, 0, num)
	case "w":
		return dt.AddDate(0, 0, num*7)
	case "ww":
		return dt.AddDate(0, 0, num*7)
	case "h":
		return dt.Add(time.Duration(num) * time.Hour)
	case "n":
		return dt.Add(time.Duration(num) * time.Minute)
	case "s":
		return dt.Add(time.Duration(num) * time.Second)
	default:
		return dt
	}
}

// vbWeekday returns VBScript-style weekday with optional first day of week
func vbWeekday(dt time.Time, firstDayOfWeek int) int {
	if firstDayOfWeek < 1 || firstDayOfWeek > 7 {
		firstDayOfWeek = 1 // vbUseSystem/vbSunday default
	}
	wd := int(dt.Weekday()) + 1 // Go: 0=Sunday; VB: 1=Sunday
	return ((wd - firstDayOfWeek + 7) % 7) + 1
}

// vbWeekNumber computes VBScript-style week number (ww) with firstweekofyear rules
func vbWeekNumber(dt time.Time, firstDayOfWeek int, firstWeekOfYear int) int {
	if firstDayOfWeek < 1 || firstDayOfWeek > 7 {
		firstDayOfWeek = 1
	}
	if firstWeekOfYear < 0 || firstWeekOfYear > 3 {
		firstWeekOfYear = 1
	}

	startOfYear := time.Date(dt.Year(), time.January, 1, 0, 0, 0, 0, dt.Location())
	firstWeekStart := firstWeekBoundary(startOfYear, firstDayOfWeek, firstWeekOfYear)
	if dt.Before(firstWeekStart) {
		// If before first week, fall back to last week of previous year
		prevYear := dt.AddDate(-1, 0, 0)
		return vbWeekNumber(prevYear, firstDayOfWeek, firstWeekOfYear)
	}
	days := int(dt.Sub(firstWeekStart).Hours() / 24)
	return (days / 7) + 1
}

// firstWeekBoundary finds the start of week 1 per VBScript rules
func firstWeekBoundary(start time.Time, firstDayOfWeek int, firstWeekOfYear int) time.Time {
	// Align to firstDayOfWeek
	shift := (int(start.Weekday()) + 1) - firstDayOfWeek
	if shift < 0 {
		shift += 7
	}
	aligned := start.AddDate(0, 0, -shift)

	switch firstWeekOfYear {
	case 2: // First week with at least four days
		for aligned.AddDate(0, 0, 3).Year() < start.Year() {
			aligned = aligned.AddDate(0, 0, 7)
		}
	case 3: // First full week
		if aligned.Year() < start.Year() {
			aligned = aligned.AddDate(0, 0, 7)
		}
	}

	return aligned
}

// vbWeekDiff calculates week difference with VBScript rules for ww interval
func vbWeekDiff(date1, date2 time.Time, firstDayOfWeek int, firstWeekOfYear int) int {
	w1 := vbWeekNumber(date1, firstDayOfWeek, firstWeekOfYear)
	w2 := vbWeekNumber(date2, firstDayOfWeek, firstWeekOfYear)
	y1 := date1.Year()
	y2 := date2.Year()
	// Adjust for cross-year scenarios
	if y1 == y2 {
		return w2 - w1
	}
	// weeks remaining in y1 plus weeks in y2
	lastDayY1 := time.Date(y1, time.December, 31, 0, 0, 0, 0, date1.Location())
	weeksInY1 := vbWeekNumber(lastDayY1, firstDayOfWeek, firstWeekOfYear)
	weeksY1Remaining := weeksInY1 - w1
	return weeksY1Remaining + w2
}

// calcDateDiff calculates difference between two dates with VBScript week rules
func calcDateDiff(interval string, date1, date2 time.Time, firstDayOfWeek int, firstWeekOfYear int) interface{} {
	if date1.After(date2) {
		date1, date2 = date2, date1 // Ensure date1 <= date2
	}

	switch interval {
	case "yyyy":
		return date2.Year() - date1.Year()
	case "q":
		years := date2.Year() - date1.Year()
		months := int(date2.Month() - date1.Month())
		return years*4 + months/3
	case "m":
		years := date2.Year() - date1.Year()
		months := int(date2.Month() - date1.Month())
		return years*12 + months
	case "y", "d":
		return int(date2.Sub(date1).Hours() / 24)
	case "w":
		return int(date2.Sub(date1).Hours() / (24 * 7))
	case "ww":
		return vbWeekDiff(date1, date2, firstDayOfWeek, firstWeekOfYear)
	case "h":
		return int(date2.Sub(date1).Hours())
	case "n":
		return int(date2.Sub(date1).Minutes())
	case "s":
		return int(date2.Sub(date1).Seconds())
	default:
		return 0
	}
}

// extractDatePart extracts a part from a date with VBScript week rules
func extractDatePart(interval string, dt time.Time, firstDayOfWeek int, firstWeekOfYear int) interface{} {
	switch interval {
	case "yyyy":
		return dt.Year()
	case "q":
		return (int(dt.Month()-1) / 3) + 1
	case "m":
		return int(dt.Month())
	case "y":
		return dt.YearDay()
	case "d":
		return dt.Day()
	case "w":
		return vbWeekday(dt, firstDayOfWeek)
	case "ww":
		return vbWeekNumber(dt, firstDayOfWeek, firstWeekOfYear)
	case "h":
		return dt.Hour()
	case "n":
		return dt.Minute()
	case "s":
		return dt.Second()
	default:
		return 0
	}
}

// formatVBDateDefault renders dates similarly to VBScript defaults
func formatVBDateDefault(dt time.Time) string {
	if dt.IsZero() {
		return ""
	}

	// Time-only values use the VBScript epoch date
	hasDate := !(dt.Year() == 1899 && dt.Month() == time.December && dt.Day() == 30)
	hasTime := dt.Hour() != 0 || dt.Minute() != 0 || dt.Second() != 0

	switch {
	case hasDate && hasTime:
		return formatDateTime(dt, FormatDateTimeGeneralDate)
	case hasDate:
		return formatDateTime(dt, FormatDateTimeShortDate)
	case hasTime:
		return formatDateTime(dt, FormatDateTimeLongTime)
	default:
		return formatDateTime(dt, FormatDateTimeShortDate)
	}
}

// formatDateTime formats a date/time according to VBScript format
func formatDateTime(dt time.Time, format int) string {
	switch format {
	case FormatDateTimeGeneralDate:
		// m/d/yy h:mm:ss AM/PM
		return dt.Format("1/2/06 3:04:05 PM")

	case FormatDateTimeLongDate:
		// dddd, mmmm dd, yyyy (e.g., Monday, January 01, 2006)
		weekday := []string{"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}[dt.Weekday()]
		month := []string{"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}[dt.Month()-1]
		return fmt.Sprintf("%s, %s %02d, %d", weekday, month, dt.Day(), dt.Year())

	case FormatDateTimeShortDate:
		// m/d/yy
		return dt.Format("1/2/06")

	case FormatDateTimeLongTime:
		// h:mm:ss AM/PM
		return dt.Format("3:04:05 PM")

	case FormatDateTimeShortTime:
		// h:mm AM/PM
		return dt.Format("3:04 PM")

	case FormatDateTimeISOWeekNumber:
		_, week := dt.ISOWeek()
		return fmt.Sprintf("%02d", week)

	case FormatDateTimeFirstWeekOfYear:
		_, week := dt.ISOWeek()
		return fmt.Sprintf("Week %d", week)

	default:
		// Default to short date
		return dt.Format("1/2/06")
	}
}
