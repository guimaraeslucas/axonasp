<%@ LANGUAGE="JScript" %>
<%
var passed = true;

// String Code Point Escapes
var s = "\u{1D306}";
if (s.length !== 2) {
	passed = false;
	Response.Write("String Code Point Escape failed: length expected 2, got " + s.length + "\n");
}

// RegExp /u flag
var re = /^\u{1D306}$/u;
if (!re.test(s)) {
	passed = false;
	Response.Write("RegExp /u flag failed to match code point escape\n");
}

var re2 = /^.$/u;
if (!re2.test(s)) {
	passed = false;
	Response.Write("RegExp /u flag failed: . should match full surrogate pair\n");
}

if (passed) {
	Response.Write("Full Unicode Support test passed!\n");
} else {
	Response.Write("Full Unicode Support test failed!\n");
}
%>