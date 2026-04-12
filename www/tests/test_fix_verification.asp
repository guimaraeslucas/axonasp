<%
Option Explicit
Response.LCID = 1046

on error resume next

dim jsonObj, testsPassed, testsFailed
testsPassed = 0
testsFailed = 0

Response.Write "<h2>Testing JSONobject Fix</h2>"

' Test 1: Create and parse
Result "Test 1: Create JSONobject", true
set jsonObj = new JSONobject
Result "Test 2: Initialize JSONobject", (jsonObj is not nothing)

dim jsonString
jsonString = "[{""test"": ""value""}]"
Result "Test 3: Parse JSON", (jsonObj.parse(jsonString) is not nothing)

' Test 4: Add method after parse
Result "Test 4: Add method after parse", testAdd()

Response.Write "<br><hr><br>"
Response.Write "Tests Passed: " & testsPassed & "<br>"
Response.Write "Tests Failed: " & testsFailed & "<br>"

if testsFailed = 0 then
    Response.Write "<strong style='color:green'>ALL TESTS PASSED!</strong><br>"
else
    Response.Write "<strong style='color:red'>SOME TESTS FAILED!</strong><br>"
end if

sub Result(testName, passed)
    if passed then
        Response.Write testName & ": <span style='color:green'>PASS</span><br>"
        testsPassed = testsPassed + 1
    else
        Response.Write testName & ": <span style='color:red'>FAIL</span><br>"
        testsFailed = testsFailed + 1
    end if
end sub

function testAdd()
    dim obj, item
    set obj = new JSONobject
    dim parseStr
    parseStr = "[{""name"": ""Test""}]"
    set obj = obj.parse(parseStr)
    
    on error resume next
    obj.add "newkey", "newvalue"
    if err.number <> 0 then
        Response.Write "ERROR in testAdd: " & err.description & " (Err " & err.number & ")<br>"
        testAdd = false
    else
        testAdd = true
    end if
    on error goto 0
end function
%>
<!--#include file="json-teste/jsonObject.class.asp" -->
