<%@ language="JScript" %>
<%
// JScript console test
console.log("Hello from JScript log");
console.info("Info message from JScript");
console.warn("Warning from JScript");
console.error("Error from JScript");

// Test with an array
var arr = [1, "two", true];
console.log(arr);

// Test with an object
var obj = { name: "AxonASP", version: 2, active: true };
console.info(obj);
%>
