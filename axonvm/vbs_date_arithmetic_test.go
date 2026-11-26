package axonvm

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Steven Borrie (@saborrie)
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

import (
	"strings"
	"testing"
)

// dateArithmeticPrelude declares the operands of the date arithmetic probes on
// saborrie/iis-oracle and a Report sub printing "value (TypeName)" or "ERR (description)".
// The probes use date literals; the operands here come from CDate of the same text so
// that the table also runs with a configured time zone other than UTC, where this
// engine currently reads a date literal's time of day as UTC (a separate issue).
const dateArithmeticPrelude = `<%
Dim d, e, p, q, a, b, h, v
d = CDate("2027-01-01 10:30:00")
e = CDate("2026-12-31")
p = CDate("1800-01-01 06:00:00")
q = CDate("1899-12-29 06:00:00")
a = CDate("1899-12-29 06:00:00")
b = CDate("1899-12-29 18:00:00")
h = CDate("2027-01-01 12:00:00")

Sub Report()
	If Err.Number <> 0 Then
		Response.Write "ERR (" & Err.Description & ")" & vbLf
	Else
		Response.Write v & " (" & TypeName(v) & ")" & vbLf
	End If
	Err.Clear
End Sub

On Error Resume Next
`

// runDateArithmeticCases evaluates each expression with "v = <expr> : Report" and
// compares the output line by line.
func runDateArithmeticCases(t *testing.T, cases [][2]string) {
	t.Helper()
	var src strings.Builder
	src.WriteString(dateArithmeticPrelude)
	for _, c := range cases {
		src.WriteString("v = " + c[0] + " : Report\n")
	}
	src.WriteString("%>")
	out := runVBSAndGetOutput(t, src.String())
	lines := strings.Split(strings.TrimSuffix(out, "\n"), "\n")
	if len(lines) != len(cases) {
		t.Fatalf("got %d lines, want %d:\n%s", len(lines), len(cases), out)
	}
	for i, c := range cases {
		if lines[i] != c[1] {
			t.Errorf("%s: got %q, want %q", c[0], lines[i], c[1])
		}
	}
}

// TestVBSDateArithmeticMatchesIIS pins arithmetic and numeric conversion on Date values
// to VBScript's day serial (days since 12/30/1899, time of day as the fraction). Every
// expected line was measured on IIS 10.0 in saborrie/iis-oracle:
//   - https://github.com/saborrie/iis-oracle/actions/runs/45592970246 (date-arithmetic.asp)
//   - https://github.com/saborrie/iis-oracle/actions/runs/45594216582 (date-arithmetic-datediff.asp)
//   - https://github.com/saborrie/iis-oracle/actions/runs/45595451489 (date-arithmetic-edges.asp)
//   - https://github.com/saborrie/iis-oracle/actions/runs/45595820636 (date-arithmetic-edges2.asp)
//
// The engine has no separate Single subtype; Long and Currency are tracked and match IIS.
// The table runs in two time zones because a Date is a local wall-clock time with no zone in VBScript.
func TestVBSDateArithmeticMatchesIIS(t *testing.T) {
	cases := [][2]string{
		// Date +/- number is a Date, in either operand order.
		{"d + 1", "1/2/2027 10:30:00 AM (Date)"},
		{"1 + d", "1/2/2027 10:30:00 AM (Date)"},
		{"d - 1", "12/31/2026 10:30:00 AM (Date)"},
		{"d + 0.5", "1/1/2027 10:30:00 PM (Date)"},
		{"1 - d", "12/28/1772 10:30:00 AM (Date)"},
		{"46389 - d", "1:30:00 PM (Date)"},
		{`d + "1"`, "1/2/2027 10:30:00 AM (Date)"},
		{`"1" + d`, "1/2/2027 10:30:00 AM (Date)"},
		{`d + "0.5"`, "1/1/2027 10:30:00 PM (Date)"},
		{`d - "1"`, "12/31/2026 10:30:00 AM (Date)"},
		{"d + True", "12/31/2026 10:30:00 AM (Date)"},
		{"d - True", "1/2/2027 10:30:00 AM (Date)"},
		{"d + Empty", "1/1/2027 10:30:00 AM (Date)"},
		{"Empty + d", "1/1/2027 10:30:00 AM (Date)"},
		{"d - Empty", "1/1/2027 10:30:00 AM (Date)"},
		{"Empty - d", "12/27/1772 10:30:00 AM (Date)"},
		{"IsNull(d + Null)", "True (Boolean)"},
		{"d + e", "1/2/2154 10:30:00 AM (Date)"},
		{"d + 1 - 1", "1/1/2027 10:30:00 AM (Date)"},
		{"d + 0.00001", "1/1/2027 10:30:01 AM (Date)"},
		// Date - Date is a Double; *, / and ^ give Doubles; \ and Mod give Longs.
		{"d - e", "1.4375 (Double)"},
		{"e - d", "-1.4375 (Double)"},
		{"d - p", "82910.6875 (Double)"},
		{"d * 1", "46388.4375 (Double)"},
		{"d / 1", "46388.4375 (Double)"},
		{"d ^ 1", "46388.4375 (Double)"},
		{"d * Empty", "0 (Double)"},
		{`d \ 1`, "46388 (Long)"},
		{"d Mod 7", "6 (Integer)"},
		{`d \ e`, "1 (Integer)"},
		{"d Mod e", "1 (Integer)"},
		{"d And 1", "0 (Integer)"},
		// A result outside the Date range (years 100 to 9999) stays a Double.
		{"d + 3000000", "3046388.4375 (Double)"},
		{"d - 1000000", "-953611.5625 (Double)"},
		// Unary minus, Int, Fix, Round and Abs keep the Date subtype.
		{"-d", "12/27/1772 10:30:00 AM (Date)"},
		{"-e", "12/28/1772 (Date)"},
		{"Int(d)", "1/1/2027 (Date)"},
		{"Fix(d)", "1/1/2027 (Date)"},
		{"Round(d)", "1/1/2027 (Date)"},
		{"Round(h)", "1/1/2027 (Date)"},
		{"Round(d, 2)", "1/1/2027 10:33:36 AM (Date)"},
		{"Abs(d)", "1/1/2027 10:30:00 AM (Date)"},
		{"Abs(-d)", "1/1/2027 10:30:00 AM (Date)"},
		{"Sgn(d)", "1 (Integer)"},
		// Conversions read the serial.
		{"CDbl(d)", "46388.4375 (Double)"},
		{`CDbl(CDate("1899-12-30"))`, "0 (Double)"},
		{"CLng(e)", "46387 (Long)"},
		{"CLng(d)", "46388 (Long)"},
		{"CLng(h)", "46388 (Long)"},
		{"CInt(d)", "ERR (Overflow)"},
		{`CInt(CDate("1950-01-01"))`, "18264 (Integer)"},
		{"CByte(d)", "ERR (Overflow)"},
		{"CSng(d) = CSng(46388.4375)", "True (Boolean)"},
		{"CCur(d) = 46388.4375", "True (Boolean)"},
		{"CBool(d)", "True (Boolean)"},
		{"CDate(46388.4375)", "1/1/2027 10:30:00 AM (Date)"},
		{"CDate(d + 1)", "1/2/2027 10:30:00 AM (Date)"},
		{"CDate(0.00001)", "12:00:01 AM (Date)"},
		{"CDate(3000000)", "ERR (Overflow)"},
		// Before 12/30/1899 the serial is negative and its fraction is still the time
		// of day, so adding a positive fraction moves the clock backwards over midnight.
		{"CDbl(p)", "-36522.25 (Double)"},
		{"CDbl(q)", "-1.25 (Double)"},
		{"CDbl(b)", "-1.75 (Double)"},
		{"p + 0.5", "1/2/1800 6:00:00 PM (Date)"},
		{"p - 0.5", "1/1/1800 6:00:00 PM (Date)"},
		{"p + 1", "1/2/1800 6:00:00 AM (Date)"},
		{"q + 1", "6:00:00 AM (Date)"},
		{"q + 2", "6:00:00 PM (Date)"},
		{"Int(p)", "12/31/1799 (Date)"},
		{"Fix(p)", "1/1/1800 (Date)"},
		{"Round(p)", "1/1/1800 (Date)"},
		{"CLng(p)", "-36522 (Long)"},
		{"-p", "12/28/1999 6:00:00 AM (Date)"},
		{"CDate(-0.25)", "6:00:00 AM (Date)"},
		{"CDate(-1.25)", "12/29/1899 6:00:00 AM (Date)"},
		{"CDate(-1.75)", "12/29/1899 6:00:00 PM (Date)"},
		// Comparisons against numbers, and between Dates, compare serials.
		{"d > 46388", "True (Boolean)"},
		{"d < 46389", "True (Boolean)"},
		{"d = 46388.4375", "True (Boolean)"},
		{"46388.4375 = d", "True (Boolean)"},
		{"d = e + 1.4375", "True (Boolean)"},
		{"d + 1 > d", "True (Boolean)"},
		{"d + 1 - 1 = d", "True (Boolean)"},
		{"p < 0", "True (Boolean)"},
		{"b < a", "True (Boolean)"},
		{"b > a", "False (Boolean)"},
		// Date functions on arithmetic results.
		{`DateAdd("d", 1, d)`, "1/2/2027 10:30:00 AM (Date)"},
		{`DateAdd("d", 1, d + 1)`, "1/3/2027 10:30:00 AM (Date)"},
		{`DateDiff("d", d, d + 1)`, "1 (Integer)"},
		{`DateDiff("h", d - 1, d)`, "24 (Integer)"},
		{"Year(d + 1)", "2027 (Integer)"},
		{"Day(d + 1)", "2 (Integer)"},
		{"Hour(d + 0.5)", "22 (Integer)"},
		{"Second(d + 0.000006)", "1 (Integer)"},
		{"FormatDateTime(d - 1, 2)", "12/31/2026 (String)"},
		// Now() and Date() results.
		{"TypeName(Date() - 1)", "Date (String)"},
		{"TypeName(Now() + 1)", "Date (String)"},
		// Without parentheses Date and Now are the same calls as Date() and Now().
		{"TypeName(Date - 1)", "Date (String)"},
		{"TypeName(Now - 1)", "Date (String)"},
		{"Int(Now() + 1) = Date() + 1", "True (Boolean)"},
		{`DateDiff("d", Date() - 1, Date())`, "1 (Integer)"},
		{`DateDiff("n", Now(), Now() + 1)`, "1440 (Integer)"},
	}
	for _, tz := range []string{"UTC", "America/Sao_Paulo", "Asia/Tokyo"} {
		t.Run(tz, func(t *testing.T) {
			withTestRuntimeTimezone(t, tz)
			runDateArithmeticCases(t, cases)
		})
	}
}

// TestVBSNowHasWholeSeconds checks Now() carries no fraction of a second, as on IIS
// (date-arithmetic-edges.asp: the serial times 86400 is whole to within float error),
// so a Now() value survives arithmetic round trips exactly.
func TestVBSNowHasWholeSeconds(t *testing.T) {
	out := runVBSAndGetOutput(t, `<%
Dim n : n = Now()
Response.Write (n + 1 - 1 = n) & "|" & (n = DateAdd("s", 0, n)) & "|" & (Abs(CDbl(n) * 86400 - Round(CDbl(n) * 86400)) < 0.0001)
%>`)
	if want := "True|True|True"; out != want {
		t.Fatalf("got %q, want %q", out, want)
	}
}

// TestVBSForLoopOverDates covers For loops whose counter or bounds are Dates
// (date-arithmetic-edges.asp and date-arithmetic-edges2.asp): they step by days and
// stop at the bound, where they used to step by one nanosecond and never finish.
func TestVBSForLoopOverDates(t *testing.T) {
	out := runVBSAndGetOutput(t, `<%
Dim d, e, x, cnt
d = #2027-01-01 10:30:00#
e = #2026-12-31#
cnt = 0
For x = e To e + 2
	cnt = cnt + 1
Next
Response.Write cnt & " " & x & " " & TypeName(x) & "|"
cnt = 0
For x = e To d Step 0.5
	cnt = cnt + 1
Next
Response.Write cnt & " " & x & " " & TypeName(x) & "|"
cnt = 0
For x = e To 46389
	cnt = cnt + 1
Next
Response.Write cnt & " " & x & " " & TypeName(x) & "|"
cnt = 0
For x = 46387 To e + 2
	cnt = cnt + 1
Next
Response.Write cnt & "|"
%>`)
	want := "3 1/3/2027 Date|3 1/1/2027 12:00:00 PM Date|3 1/3/2027 Date|3|"
	if out != want {
		t.Fatalf("got  %q\nwant %q", out, want)
	}
}

// TestVBSDateArithmeticKeepsWallClockAcrossDST checks the serial is taken from, and
// turned back into, local wall-clock time: adding half a day to midnight on the day
// clocks go forward gives noon, and Date + 0 is the same Date in the afternoon.
func TestVBSDateArithmeticKeepsWallClockAcrossDST(t *testing.T) {
	withTestRuntimeTimezone(t, "America/New_York")
	out := runVBSAndGetOutput(t, `<%
Dim m, f
m = CDate("2027-03-14")
f = CDate("2027-03-14 15:00:00")
Response.Write (m + 0.5) & "|" & (f + 0) & "|" & (f - m) & "|" & CDate(CDbl(f)) & "|" & (m + 1 - m)
%>`)
	want := "3/14/2027 12:00:00 PM|3/14/2027 3:00:00 PM|0.625|3/14/2027 3:00:00 PM|1"
	if out != want {
		t.Fatalf("got  %q\nwant %q", out, want)
	}
}
