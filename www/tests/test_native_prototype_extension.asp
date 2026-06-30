<%@ Language="JScript" %><%
/*
 * AxonASP Native Prototype Extension Test
 *
 * Verifies that user-defined methods and properties can be added
 * to built-in prototype objects (Number, String, Boolean) and
 * accessed from primitive values.
 */

Response.Clear();
Response.Status         = 200;
Response.ContentType    = "text/plain";
Response.CharSet        = "utf-8";
Response.CacheControl   = "max-age=0, no-cache, no-store";

function prefixPad(intLength, strPadChar) {
    if (typeof intLength !== "number") {
        var intLength = 2;
    }
    if (typeof strPadChar !== "string") {
        var strPadChar = "0";
    }

    var strInput = String(this);

    for (var i = strInput.length; i < intLength; ++i) {
        strInput = strPadChar + strInput;
    }

    return strInput;
}

// Assign to both Number and String prototypes
Number.prototype.prefixPad = prefixPad;
String.prototype.prefixPad = prefixPad;

// Boolean prototype extension
Boolean.prototype.toFlag = function() {
    return this ? "TRUE" : "FALSE";
};

// Write results
Response.Write("Number.prefixPad: " + ((7).prefixPad()) + "\n");
Response.Write("Number.prefixPad: " + ((9).prefixPad(3, "9")) + "\n");
Response.Write("String.prefixPad: " + ("07".prefixPad(3)) + "\n");
Response.Write("String.prefixPad: " + ("H".prefixPad(5, "O")) + "\n");

// Boolean tests
Response.Write("Boolean.toFlag: " + (true.toFlag()) + "\n");
Response.Write("Boolean.toFlag: " + (false.toFlag()) + "\n");

// Property access on primitives
Number.prototype.customData = "numData";
Response.Write("Number.customData: " + ((100).customData) + "\n");

String.prototype.customData = "strData";
Response.Write("String.customData: " + ("hello".customData) + "\n");

Response.Write("--- All native prototype extension tests completed. ---");
%>