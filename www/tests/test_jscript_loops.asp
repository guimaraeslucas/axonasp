<%
' Test JScript loop compilation and execution
Set testSuite = Server.CreateObject("G3TestSuite")

' Test 1: While loop
testSuite.BeginTest "JScript While Loop"
testSuite.SetVar "whileResult", ""
%>
<script language="jscript" runat="server">
    var i = 0;
    var result = "";
    while (i < 3) {
        result += i + ",";
        i++;
    }
    testSuite.SetVar("whileResult", result);
</script>
<%
expected = "0,1,2,"
actual = testSuite.GetVar("whileResult")
testSuite.AssertEquals actual, expected, "While loop should iterate 0,1,2"
testSuite.EndTest

' Test 2: Do-While loop
testSuite.BeginTest "JScript Do-While Loop"
testSuite.SetVar "doWhileResult", ""
%>
<script language="jscript" runat="server">
    var j = 0;
    var result = "";
    do {
        result += j + ",";
        j++;
    } while (j < 2);
    testSuite.SetVar("doWhileResult", result);
</script>
<%
expected = "0,1,"
actual = testSuite.GetVar("doWhileResult")
testSuite.AssertEquals actual, expected, "Do-while loop should iterate 0,1"
testSuite.EndTest

' Test 3: For loop
testSuite.BeginTest "JScript For Loop"
testSuite.SetVar "forResult", ""
%>
<script language="jscript" runat="server">
    var result = "";
    for (var k = 0; k < 3; k++) {
        result += k + ",";
    }
    testSuite.SetVar("forResult", result);
</script>
<%
expected = "0,1,2,"
actual = testSuite.GetVar("forResult")
testSuite.AssertEquals actual, expected, "For loop should iterate 0,1,2"
testSuite.EndTest

' Test 4: Break statement
testSuite.BeginTest "JScript Break Statement"
testSuite.SetVar "breakResult", ""
%>
<script language="jscript" runat="server">
    var result = "";
    for (var m = 0; m < 5; m++) {
        if (m === 3) break;
        result += m + ",";
    }
    testSuite.SetVar("breakResult", result);
</script>
<%
expected = "0,1,2,"
actual = testSuite.GetVar("breakResult")
testSuite.AssertEquals actual, expected, "Break should exit at m=3"
testSuite.EndTest

' Test 5: Continue statement
testSuite.BeginTest "JScript Continue Statement"
testSuite.SetVar "continueResult", ""
%>
<script language="jscript" runat="server">
    var result = "";
    for (var n = 0; n < 5; n++) {
        if (n === 2) continue;
        result += n + ",";
    }
    testSuite.SetVar("continueResult", result);
</script>
<%
expected = "0,1,3,4,"
actual = testSuite.GetVar("continueResult")
testSuite.AssertEquals actual, expected, "Continue should skip n=2"
testSuite.EndTest

' Test 6: Arithmetic operators
testSuite.BeginTest "JScript Arithmetic Operators"
%>
<script language="jscript" runat="server">
    testSuite.SetVar(
        "arithmetic",
        5 - 2 + "," + 3 * 4 + "," + 10 / 2 + "," + (10 % 3)
    );
</script>
<%
expected = "3,12,5,1"
actual = testSuite.GetVar("arithmetic")
testSuite.AssertEquals actual, expected, "Arithmetic operators should work correctly"
testSuite.EndTest

' Test 7: Comparison operators
testSuite.BeginTest "JScript Comparison Operators"
%>
<script language="jscript" runat="server">
    testSuite.SetVar(
        "comparisons",
        (5 > 3 ? "true" : "false") +
            "," +
            (2 < 3 ? "true" : "false") +
            "," +
            (5 >= 5 ? "true" : "false") +
            "," +
            (3 <= 2 ? "true" : "false")
    );
</script>
<%
expected = "true,true,true,false"
actual = testSuite.GetVar("comparisons")
testSuite.AssertEquals actual, expected, "Comparison operators should work correctly"
testSuite.EndTest

' Test 8: Logical operators
testSuite.BeginTest "JScript Logical Operators"
%>
<script language="jscript" runat="server">
    testSuite.SetVar(
        "logical",
        (true && true ? "true" : "false") +
            "," +
            (false || true ? "true" : "false") +
            "," +
            (!true ? "true" : "false")
    );
</script>
<%
expected = "true,true,false"
actual = testSuite.GetVar("logical")
testSuite.AssertEquals actual, expected, "Logical operators should work correctly"
testSuite.EndTest

testSuite.Summary
%>
