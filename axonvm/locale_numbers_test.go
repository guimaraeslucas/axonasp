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

import (
	"strconv"
	"strings"
	"testing"

	"g3pix.com.br/axonasp/v2/vbscript"
)

// TestVBScriptLocaleNumbersMatchIIS verifies locale-aware number formatting and parsing
// across SetLocale / Session.LCID 1031 (de-DE), 1036 (fr-FR), 2057 (en-GB), and 1033 (en-US).
func TestVBScriptLocaleNumbersMatchIIS(t *testing.T) {
	t.Run("SetLocale_1031_CStr_Concat_CCur_CDbl", func(t *testing.T) {
		src := `<%
On Error Resume Next
SetLocale 1031
Dim a, b, c, d, e, f, g, h
a = CStr(1234.5)
b = 1234.5 & ""
c = CStr(CCur(1234.5))
d = CStr(CCur("15.00"))
e = VarType(CCur("15.00"))
f = CStr(CDbl("150.00"))
g = CStr(CDbl("1234,5"))
h = CStr(CDbl("1.234,5"))
Response.Write a & "|" & b & "|" & c & "|" & d & "|" & e & "|" & f & "|" & g & "|" & h
%>`
		want := "1234,5|1234,5|1234,5|1500|6|15000|1234,5|1234,5"
		if got := runVBSAndGetOutput(t, src); got != want {
			t.Fatalf("got %q, want %q", got, want)
		}
	})

	t.Run("Session_LCID_1031_ResponseWrite_And_QueryString", func(t *testing.T) {
		src := `<%
Session.LCID = 1031
Dim x
x = 1234.5
Response.Write "&amount=" & x & "|"
%><%= x %>|<%= 1234.5 %>`
		want := "&amount=1234,5|1234,5|1234,5"
		if got := runVBSAndGetOutput(t, src); got != want {
			t.Fatalf("got %q, want %q", got, want)
		}
	})

	t.Run("SetLocale_1036_French", func(t *testing.T) {
		src := `<%
On Error Resume Next
SetLocale 1036
Dim a, b, c, errDot
a = CStr(1234.5)
b = 1234.5 & ""
c = CStr(CDbl("1 234,5"))
Err.Clear
Dim bad
bad = CDbl("150.00")
errDot = Err.Number
Response.Write a & "|" & b & "|" & c & "|" & errDot
%>`
		typeMismatch := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.TypeMismatch))
		want := "1234,5|1234,5|1234,5|" + typeMismatch
		if got := runVBSAndGetOutput(t, src); got != want {
			t.Fatalf("got %q, want %q", got, want)
		}
	})

	t.Run("SetLocale_2057_EnglishUK", func(t *testing.T) {
		src := `<%
SetLocale 2057
Response.Write CStr(1234.5) & "|" & (1234.5 & "") & "|" & CStr(CCur(1234.5)) & "|" & CStr(CCur("15.00")) & "|" & CStr(CDbl("150.00"))
%>`
		want := "1234.5|1234.5|1234.5|15|150"
		if got := runVBSAndGetOutput(t, src); got != want {
			t.Fatalf("got %q, want %q", got, want)
		}
	})

	t.Run("Locale_1031_Arithmetic_And_StringBuiltins", func(t *testing.T) {
		src := `<%
SetLocale 1031
Dim sAdd, sNeg
sAdd = "12,5" + 1
sNeg = -"12,5"
Response.Write TypeName(sAdd) & ":" & CStr(sAdd) & "|" & TypeName(sNeg) & ":" & CStr(sNeg) & "|" & Replace(1234.5, ",", ".") & "|" & InStr(1234.5, ",") & "|" & IsNumeric("12,5")
%>`
		want := "Double:13,5|Double:-12,5|1234.5|5|True"
		if got := runVBSAndGetOutput(t, src); got != want {
			t.Fatalf("got %q, want %q", got, want)
		}
	})
}

// TestVBScriptRuntimeSemanticsFixes verifies Currency subtype, CLng TypeName,
// Asc("") error, AscW/ChrW signed surrogate handling, and typed Dictionary keys.
func TestVBScriptRuntimeSemanticsFixes(t *testing.T) {
	t.Run("CCur_Currency_Subtype_And_Four_Places", func(t *testing.T) {
		src := `<%
Dim c1, c2, cSum
c1 = CCur(1234.5)
c2 = CCur(1.23456)
cSum = CCur(0.1) + CCur(0.2)
Response.Write VarType(c1) & "|" & TypeName(c1) & "|" & CStr(c2) & "|" & VarType(cSum) & "|" & CStr(cSum)
%>`
		want := "6|Currency|1.2346|6|0.3"
		if got := runVBSAndGetOutput(t, src); got != want {
			t.Fatalf("got %q, want %q", got, want)
		}
	})

	t.Run("CLng_Preserves_Long_TypeName_And_VarType", func(t *testing.T) {
		src := `<%
Dim s, l, i
s = 5
l = CLng(5)
i = CInt(l)
Response.Write TypeName(s) & ":" & VarType(s) & "|" & TypeName(l) & ":" & VarType(l) & "|" & TypeName(i) & ":" & VarType(i)
%>`
		want := "Integer:2|Long:3|Integer:2"
		if got := runVBSAndGetOutput(t, src); got != want {
			t.Fatalf("got %q, want %q", got, want)
		}
	})

	t.Run("AscW_ChrW_Signed_And_Surrogates", func(t *testing.T) {
		src := `<%
Response.Write AscW(ChrW(55357)) & "|" & AscW(ChrW(65535)) & "|" & AscW(ChrW(32768)) & "|" & AscW(ChrW(-1)) & "|" & Len(ChrW(55357))
%>`
		want := "-10179|-1|-32768|-1|1"
		if got := runVBSAndGetOutput(t, src); got != want {
			t.Fatalf("got %q, want %q", got, want)
		}
	})

	t.Run("Asc_Empty_Raises_Error_5", func(t *testing.T) {
		src := `<%
On Error Resume Next
Err.Clear
Dim a, b, e1, e2
a = Asc("")
e1 = Err.Number
Err.Clear
b = AscW("")
e2 = Err.Number
Response.Write e1 & "|" & e2
%>`
		invalidProc := strconv.Itoa(vbscript.HRESULTFromVBScriptCode(vbscript.InvalidProcedureCallOrArgument))
		want := invalidProc + "|" + invalidProc
		if got := runVBSAndGetOutput(t, src); got != want {
			t.Fatalf("got %q, want %q", got, want)
		}
	})

	t.Run("Dictionary_Keys_Preserve_Type", func(t *testing.T) {
		src := `<%
Dim d, k
Set d = Server.CreateObject("Scripting.Dictionary")
d.Add 1, "num"
Dim existsStr, existsInt, existsLong
existsStr = d.Exists("1")
existsInt = d.Exists(1)
existsLong = d.Exists(CLng(1))
d.Add "1", "str"
k = d.Keys()
Response.Write existsStr & "|" & existsInt & "|" & existsLong & "|" & d.Count & "|" & d(1) & "|" & d("1") & "|" & TypeName(k(0)) & "|" & TypeName(k(1))
%>`
		want := "False|True|True|2|num|str|Integer|String"
		if got := runVBSAndGetOutput(t, src); got != want {
			t.Fatalf("got %q, want %q", got, want)
		}
	})

	t.Run("JScript_Number_Formatting_Unaffected_By_LCID_1031", func(t *testing.T) {
		src := `<%@ Language="JScript" %><%
Session.LCID = 1031;
var x = 1234.5;
Response.Write(String(x) + "|" + encodeURIComponent(x) + "|" + JSON.stringify({ v: x }));
%>`
		want := `1234.5|1234.5|{"v":1234.5}`
		got := runVBSAndGetOutput(t, src)
		if !strings.Contains(got, want) || got != want {
			t.Fatalf("got %q, want %q", got, want)
		}
	})
}
