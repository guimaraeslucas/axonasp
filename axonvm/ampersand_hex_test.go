package axonvm

import "testing"

// TestAmpersandHexPrefixRequiresAdjacency pins the boundary between a hex/octal
// literal and concatenation with an identifier that merely looks like one.
//
// `&hB` is the literal 11. `& hb`, with whitespace, is concatenation with the
// variable hb — verified against IIS 10.0 / VBScript 5.8.16384, where
// `Dim hb : hb = "VALUE"` makes `"[" & hb & "]TAIL"` render `[VALUE]TAIL`.
//
// Treating the whitespace form as a literal is not merely a wrong value: the
// lexer consumes the identifier and the remainder of the expression silently
// disappears, so `"[" & hb & "]TAIL"` rendered `[`. Any identifier matching
// [hH][0-9a-fA-F]+ was affected — hb, h1, hf, h2a — while names that are not
// wholly hex digits (hz, hello) were not.
func TestAmpersandHexPrefixRequiresAdjacency(t *testing.T) {
	tests := []struct {
		name   string
		source string
		want   string
	}{
		// The regression: an identifier after `& ` must stay an identifier.
		{
			name:   "identifier after ampersand and space",
			source: `<% Dim hb : hb = "VALUE" : Response.Write "[" & hb & "]TAIL" %>`,
			want:   "[VALUE]TAIL",
		},
		{
			name:   "single hex digit identifier",
			source: `<% Dim h1 : h1 = "V" : Response.Write "[" & h1 & "]TAIL" %>`,
			want:   "[V]TAIL",
		},
		{
			name:   "identifier that is not wholly hex digits",
			source: `<% Dim hz : hz = "V" : Response.Write "[" & hz & "]TAIL" %>`,
			want:   "[V]TAIL",
		},
		{
			name:   "parenthesised identifier",
			source: `<% Dim hb : hb = "V" : Response.Write "[" & (hb) & "]TAIL" %>`,
			want:   "[V]TAIL",
		},
		{
			name:   "member access on an object named like a hex literal",
			source: `<% Dim hb : Set hb = Server.CreateObject("Scripting.Dictionary") : hb.Add "k", 1 : Response.Write "[" & hb.Count & "]TAIL" %>`,
			want:   "[1]TAIL",
		},

		// Literals must keep working: the prefix is adjacent to the ampersand.
		{
			name:   "adjacent hex literal",
			source: `<% Response.Write CStr(&hB) %>`,
			want:   "11",
		},
		{
			name:   "adjacent hex literal, multiple digits",
			source: `<% Response.Write CStr(&hFF) & "," & CStr(&hFF00) %>`,
			want:   "255,65280",
		},
		{
			name:   "adjacent octal literal",
			source: `<% Response.Write CStr(&o17) %>`,
			want:   "15",
		},
		{
			name:   "hex literal concatenated with a string",
			source: `<% Response.Write "[" & CStr(&hB) & "]TAIL" %>`,
			want:   "[11]TAIL",
		},
		{
			name:   "builtin whose name begins with a hex-digit run",
			source: `<% Response.Write "[" & Hex(255) & "]TAIL" %>`,
			want:   "[FF]TAIL",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := runVBSAndGetOutput(t, tt.source); got != tt.want {
				t.Fatalf("%s\n  got  %q\n  want %q", tt.source, got, tt.want)
			}
		})
	}
}
