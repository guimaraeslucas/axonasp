/*
 * AxonASP Server
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
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 */
package vbscript

import (
	"testing"
	"time"
)

// TestGetDateUSFormat verifies that date literals (#...#) are always interpreted
// as US date format (M/D/YYYY), regardless of any locale setting.
// This is a critical VBScript compatibility requirement.
func TestGetDateUSFormat(t *testing.T) {
	tests := []struct {
		input  string
		year   int
		month  time.Month
		day    int
		hour   int
		minute int
		second int
	}{
		// US format M/D/YYYY
		{input: "3/4/2026", year: 2026, month: time.March, day: 4},
		{input: "1/2/2023", year: 2023, month: time.January, day: 2},
		{input: "12/31/2025", year: 2025, month: time.December, day: 31},
		{input: "7/4/1776", year: 1776, month: time.July, day: 4},

		// US format with leading zeros
		{input: "03/04/2026", year: 2026, month: time.March, day: 4},
		{input: "01/02/2023", year: 2023, month: time.January, day: 2},

		// US format with spaces around delimiters (VBScript allows this)
		{input: "3 / 4 / 2026", year: 2026, month: time.March, day: 4},
		{input: " 3/4/2026 ", year: 2026, month: time.March, day: 4},

		// ISO format YYYY-M-D
		{input: "2026-3-4", year: 2026, month: time.March, day: 4},
		{input: "2026-03-04", year: 2026, month: time.March, day: 4},
		{input: "2023-01-02", year: 2023, month: time.January, day: 2},

		// US format with time (12h AM/PM)
		{input: "3/4/2026 10:30:00 AM", year: 2026, month: time.March, day: 4, hour: 10, minute: 30, second: 0},
		{input: "12/31/2025 11:59:59 PM", year: 2025, month: time.December, day: 31, hour: 23, minute: 59, second: 59},

		// US format with time (24h)
		{input: "3/4/2026 15:04:05", year: 2026, month: time.March, day: 4, hour: 15, minute: 4, second: 5},

		// ISO format with time
		{input: "2026-03-04 15:04:05", year: 2026, month: time.March, day: 4, hour: 15, minute: 4, second: 5},

		// Time only (VBScript serial 0 date component is December 30, 1899)
		{input: "3:04:05 PM", year: 1899, month: time.December, day: 30, hour: 15, minute: 4, second: 5},
		{input: "15:04:05", year: 1899, month: time.December, day: 30, hour: 15, minute: 4, second: 5},
		{input: "13:45:30", year: 1899, month: time.December, day: 30, hour: 13, minute: 45, second: 30},
		{input: "1:45:30 PM", year: 1899, month: time.December, day: 30, hour: 13, minute: 45, second: 30},
		{input: "1:45:30PM", year: 1899, month: time.December, day: 30, hour: 13, minute: 45, second: 30},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			result, err := GetDate(tt.input)
			if err != nil {
				t.Fatalf("GetDate(%q) returned error: %v", tt.input, err)
			}
			if result.Year() != tt.year {
				t.Errorf("GetDate(%q): expected year %d, got %d", tt.input, tt.year, result.Year())
			}
			if result.Month() != tt.month {
				t.Errorf("GetDate(%q): expected month %s, got %s", tt.input, tt.month, result.Month())
			}
			if result.Day() != tt.day {
				t.Errorf("GetDate(%q): expected day %d, got %d", tt.input, tt.day, result.Day())
			}
			if tt.hour != 0 || tt.minute != 0 || tt.second != 0 {
				if result.Hour() != tt.hour {
					t.Errorf("GetDate(%q): expected hour %d, got %d", tt.input, tt.hour, result.Hour())
				}
				if result.Minute() != tt.minute {
					t.Errorf("GetDate(%q): expected minute %d, got %d", tt.input, tt.minute, result.Minute())
				}
				if result.Second() != tt.second {
					t.Errorf("GetDate(%q): expected second %d, got %d", tt.input, tt.second, result.Second())
				}
			}
		})
	}
}

// TestGetDateRejectsNonUSFormat verifies that date literals in non-US formats
// (such as DD/MM/YYYY) are rejected. VBScript date literals MUST be US or ISO format.
func TestGetDateRejectsNonUSFormat(t *testing.T) {
	// These are DD/MM/YYYY format dates that could be ambiguous.
	// In US format, the first number is the month, so 13/4/2026 is invalid.
	// 4/13/2026 is valid US (April 13), but GetDate should still parse it as US.
	invalidUS := []string{
		"13/4/2026",  // month 13 is invalid
		"13/04/2026", // month 13 is invalid
	}

	for _, input := range invalidUS {
		t.Run(input, func(t *testing.T) {
			_, err := GetDate(input)
			if err == nil {
				t.Errorf("GetDate(%q): expected error for non-US date, but got none", input)
			}
		})
	}

	// Verify that 4/13/2026 is parsed as US (April 13, not day 4 month 13)
	t.Run("4/13/2026 must be April", func(t *testing.T) {
		result, err := GetDate("4/13/2026")
		if err != nil {
			t.Fatalf("GetDate(4/13/2026) returned error: %v", err)
		}
		if result.Month() != time.April || result.Day() != 13 {
			t.Errorf("GetDate(4/13/2026): expected April 13 (US), got %s %d", result.Month(), result.Day())
		}
	})
}

// TestGetDateISOFormat verifies ISO 8601 date literals are supported.
func TestGetDateISOFormat(t *testing.T) {
	tests := []struct {
		input string
		year  int
		month time.Month
		day   int
	}{
		{input: "2026-03-04", year: 2026, month: time.March, day: 4},
		{input: "2026-3-4", year: 2026, month: time.March, day: 4},
		{input: "2023-01-02", year: 2023, month: time.January, day: 2},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			result, err := GetDate(tt.input)
			if err != nil {
				t.Fatalf("GetDate(%q) returned error: %v", tt.input, err)
			}
			if result.Year() != tt.year || result.Month() != tt.month || result.Day() != tt.day {
				t.Errorf("GetDate(%q): expected %d-%s-%d, got %d-%s-%d",
					tt.input, tt.year, tt.month, tt.day,
					result.Year(), result.Month(), result.Day())
			}
		})
	}
}

// TestGetDateEmptyAndInvalid verifies error handling for invalid date literals.
func TestGetDateEmptyAndInvalid(t *testing.T) {
	invalid := []string{
		"",
		"not a date",
		"abc-def-ghi",
		"13/13/2026", // month 13, day 13
	}

	for _, input := range invalid {
		t.Run(input, func(t *testing.T) {
			_, err := GetDate(input)
			if err == nil {
				t.Errorf("GetDate(%q): expected error, got none", input)
			}
		})
	}
}
