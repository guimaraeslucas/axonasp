<%@ LANGUAGE="JScript" %>
<%
var passed = true;

try {
	let a = 1;
	{
		let b = a;
		let a = 2; // TDZ error
	}
	passed = false;
} catch (e) {
	if (e.message.indexOf("Cannot access 'a' before initialization") !== -1) {
		// Expected
	} else {
		passed = false;
		Response.Write("Unexpected error: " + e.message + "\n");
	}
}

try {
	let x = 1;
	{
		let y = x;
		const x = 2; // TDZ error
	}
	passed = false;
} catch (e) {
	if (e.message.indexOf("Cannot access 'x' before initialization") !== -1) {
		// Expected
	} else {
		passed = false;
		Response.Write("Unexpected error: " + e.message + "\n");
	}
}

if (passed) {
	Response.Write("TDZ let/const test passed!\n");
} else {
	Response.Write("TDZ let/const test failed!\n");
}
%>