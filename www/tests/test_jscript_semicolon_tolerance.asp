<%@ Language="JScript" %>
<%
function testTryCatch() {
    try {
        throw "boom";
    };
    catch (e) {
        Response.Write("caught: " + e + "\n");
    };
}

function testIfElse(n) {
    if (n == 1) {
        return "one";
    };
    else if (n == 2) {
        return "two";
    };
    else {
        return "other";
    };
}

testTryCatch();
Response.Write("Test If 1: " + testIfElse(1) + "\n");
Response.Write("Test If 2: " + testIfElse(2) + "\n");
Response.Write("Test If 9: " + testIfElse(9) + "\n");
%>
