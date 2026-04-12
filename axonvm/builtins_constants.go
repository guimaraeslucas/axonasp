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
 * Attribution Notice:
 * If this software is used in other projects, the name "AxonASP Server"
 * must be cited in the documentation or "About" section.
 *
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 */
package axonvm

// VBSConstant holds a predefined VBScript global constant — name and its immutable value.
// Constants are injected into the compiler globals and VM globals in declaration order,
// immediately after the built-in function slots, so indices are deterministic and stable.
type VBSConstant struct {
	Name string
	Val  Value
}

// VBSConstants is the ordered registry of all predefined VBScript global constants.
// Order MUST match between this slice and the compiler/VM injection loops.
// Source of truth: Microsoft VBScript Language Reference (5.8).
var VBSConstants = []VBSConstant{
	// ── String / Character constants ──────────────────────────────
	{"vbCr", NewString("\r")},            // Carriage-return character (Chr 13)
	{"vbLf", NewString("\n")},            // Linefeed character (Chr 10)
	{"vbCrLf", NewString("\r\n")},        // CR+LF combination
	{"vbNewLine", NewString("\r\n")},     // System-defined newline (same as vbCrLf on Windows)
	{"vbNullChar", NewString("\x00")},    // Null character (Chr 0)
	{"vbNullString", NewString("")},      // Null-length string (distinct from Empty in COM but same value here)
	{"vbTab", NewString("\t")},           // Horizontal tab (Chr 9)
	{"vbBack", NewString("\x08")},        // Backspace character (Chr 8)
	{"vbFormFeed", NewString("\x0C")},    // Form-feed character (Chr 12)
	{"vbVerticalTab", NewString("\x0B")}, // Vertical-tab character (Chr 11)

	// ── Boolean constants ─────────────────────────────────────────
	// VBScript True = -1, False = 0 (stored as integers per spec)
	{"vbTrue", NewInteger(-1)},
	{"vbFalse", NewInteger(0)},

	// ── Variant type constants (used with VarType / TypeName) ─────
	{"vbEmpty", NewInteger(0)},
	{"vbNull", NewInteger(1)},
	{"vbInteger", NewInteger(2)},
	{"vbLong", NewInteger(3)},
	{"vbSingle", NewInteger(4)},
	{"vbDouble", NewInteger(5)},
	{"vbCurrency", NewInteger(6)},
	{"vbDate", NewInteger(7)},
	{"vbString", NewInteger(8)},
	{"vbObject", NewInteger(9)},
	{"vbError", NewInteger(10)},
	{"vbBoolean", NewInteger(11)},
	{"vbVariant", NewInteger(12)},
	{"vbDataObject", NewInteger(13)},
	{"vbDecimal", NewInteger(14)},
	{"vbByte", NewInteger(17)},
	{"vbArray", NewInteger(8192)},

	// ── FormatDateTime / date-format constants ────────────────────
	{"vbGeneralDate", NewInteger(0)}, // Display date and time per locale
	{"vbLongDate", NewInteger(1)},    // Display date using locale long-date format
	{"vbShortDate", NewInteger(2)},   // Display date using locale short-date format
	{"vbLongTime", NewInteger(3)},    // Display time using locale long-time format
	{"vbShortTime", NewInteger(4)},   // Display time using 24-hour format (HH:MM)

	// ── String comparison constants ───────────────────────────────
	{"vbBinaryCompare", NewInteger(0)},   // Binary (case-sensitive) comparison
	{"vbTextCompare", NewInteger(1)},     // Text (case-insensitive) comparison
	{"vbDatabaseCompare", NewInteger(2)}, // Database comparison (locale-aware)

	// ── StrConv constants ─────────────────────────────────────────
	{"vbUpperCase", NewInteger(1)},     // Convert string to upper-case
	{"vbLowerCase", NewInteger(2)},     // Convert string to lower-case
	{"vbProperCase", NewInteger(3)},    // Convert first letter of each word to upper-case
	{"vbWide", NewInteger(4)},          // Convert narrow (single-byte) to wide (double-byte)
	{"vbNarrow", NewInteger(8)},        // Convert wide to narrow
	{"vbKatakana", NewInteger(16)},     // Convert Hiragana to Katakana
	{"vbHiragana", NewInteger(32)},     // Convert Katakana to Hiragana
	{"vbUnicode", NewInteger(64)},      // Convert from default code page to Unicode
	{"vbFromUnicode", NewInteger(128)}, // Convert from Unicode to default code page

	// ── Calendar / day-of-week constants ─────────────────────────
	{"vbSunday", NewInteger(1)},
	{"vbMonday", NewInteger(2)},
	{"vbTuesday", NewInteger(3)},
	{"vbWednesday", NewInteger(4)},
	{"vbThursday", NewInteger(5)},
	{"vbFriday", NewInteger(6)},
	{"vbSaturday", NewInteger(7)},
	{"vbJanuary", NewInteger(1)},
	{"vbFebruary", NewInteger(2)},
	{"vbMarch", NewInteger(3)},
	{"vbApril", NewInteger(4)},
	{"vbMay", NewInteger(5)},
	{"vbJune", NewInteger(6)},
	{"vbJuly", NewInteger(7)},
	{"vbAugust", NewInteger(8)},
	{"vbSeptember", NewInteger(9)},
	{"vbOctober", NewInteger(10)},
	{"vbNovember", NewInteger(11)},
	{"vbDecember", NewInteger(12)},

	// First-day-of-week constants (used by DateDiff, WeekDay etc.)
	{"vbUseSystemDayOfWeek", NewInteger(0)}, // Use system regional setting
	{"vbUseSystem", NewInteger(0)},          // Use system locale defaults
	{"vbFirstJan1", NewInteger(1)},          // Start with week that contains Jan 1
	{"vbFirstFourDays", NewInteger(2)},      // Start with week that has at least 4 days in new year
	{"vbFirstFullWeek", NewInteger(3)},      // Start with first full week of new year

	// ── DatePart / DateAdd interval constants (informational) ─────
	// These are passed as strings in VBScript; no numeric constants exist.
	// Defined here only for documentation completeness — not used as numeric slots.

	// ── Color constants ───────────────────────────────────────────
	{"vbBlack", NewInteger(0x000000)},
	{"vbRed", NewInteger(0xFF0000)},
	{"vbGreen", NewInteger(0x00FF00)},
	{"vbYellow", NewInteger(0xFFFF00)},
	{"vbBlue", NewInteger(0x0000FF)},
	{"vbMagenta", NewInteger(0xFF00FF)},
	{"vbCyan", NewInteger(0x00FFFF)},
	{"vbWhite", NewInteger(0xFFFFFF)},

	// ── MsgBox button/return constants (rarely used in ASP but present in spec) ─
	{"vbOK", NewInteger(1)},
	{"vbCancel", NewInteger(2)},
	{"vbAbort", NewInteger(3)},
	{"vbRetry", NewInteger(4)},
	{"vbIgnore", NewInteger(5)},
	{"vbYes", NewInteger(6)},
	{"vbNo", NewInteger(7)},
	{"vbOKOnly", NewInteger(0)},
	{"vbOKCancel", NewInteger(1)},
	{"vbAbortRetryIgnore", NewInteger(2)},
	{"vbYesNoCancel", NewInteger(3)},
	{"vbYesNo", NewInteger(4)},
	{"vbRetryCancel", NewInteger(5)},
	{"vbCritical", NewInteger(16)},
	{"vbQuestion", NewInteger(32)},
	{"vbExclamation", NewInteger(48)},
	{"vbInformation", NewInteger(64)},
	{"vbDefaultButton1", NewInteger(0)},
	{"vbDefaultButton2", NewInteger(256)},
	{"vbDefaultButton3", NewInteger(512)},
	{"vbDefaultButton4", NewInteger(768)},
	{"vbApplicationModal", NewInteger(0)},
	{"vbSystemModal", NewInteger(4096)},
	{"vbMsgBoxHelpButton", NewInteger(16384)},
	{"vbMsgBoxSetForeground", NewInteger(65536)},
	{"vbMsgBoxRight", NewInteger(524288)},
	{"vbMsgBoxRtlReading", NewInteger(1048576)},
	{"vbObjectError", NewInteger(-2147221504)},
	{"vbASCII", NewInteger(0)},

	// ── Tristate constants (used by some format functions) ────────
	{"vbUseDefault", NewInteger(-2)}, // Use setting from regional settings

	// ── Math constants (not standard VBScript but widely expected) ─
	// none — Pi/E are available via Atn(1)*4 / Exp(1) per VBScript convention
}
