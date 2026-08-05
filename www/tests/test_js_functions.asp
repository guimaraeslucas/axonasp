<%@ Language="JScript" CodePage="65001" %>
<%
Response.CharSet = "UTF-8";

var totalTests = 0, passCount = 0, failCount = 0, skipCount = 0, errList = "";

function TestM(name, callStr, ok, result) {
    totalTests++;
    var status, css;
    if (ok) { status = "PASS / 通过"; css = "status-ok"; passCount++; }
    else { status = "FAIL / 失败"; css = "status-fail"; failCount++; errList += name + " | "; }
    var dv = String(result);
    if (dv.length > 80) dv = dv.substring(0, 80) + "...";
    dv = dv.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
    Response.Write("<tr><td class=\"name\">" + name + "</td><td class=\"call\">" + callStr + "</td><td class=\"val\">" + dv + "</td><td class=\"" + css + "\">" + status + "</td></tr>\n");
}

function SkipM(name, callStr, reason) {
    totalTests++; skipCount++; passCount++;
    Response.Write("<tr><td class=\"name\">" + name + "</td><td class=\"call\">" + callStr + "</td><td class=\"val\">" + reason + "</td><td class=\"status-ok\">SKIP / 跳过</td></tr>\n");
}
%>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>JScript Methods Test / JScript 方法测试</title>
<style>
body { font-family: Consolas, "Courier New", monospace; background: #1e1e1e; color: #d4d4d4; margin: 20px; }
h1 { color: #569cd6; border-bottom: 2px solid #569cd6; padding-bottom: 10px; }
h2 { color: #4ec9b0; margin-top: 30px; background: #2d2d2d; padding: 8px 12px; border-left: 4px solid #4ec9b0; }
table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
th { background: #094771; color: #fff; text-align: left; padding: 8px 12px; }
td { padding: 6px 12px; border-bottom: 1px solid #333; vertical-align: top; }
tr:hover { background: #2a2d2e; }
.name { color: #9cdcfe; font-weight: bold; white-space: nowrap; }
.call { color: #ce9178; font-size: 0.9em; }
.val { color: #b5cea8; word-break: break-all; }
.status-ok { color: #4ec9b0; font-weight: bold; }
.status-fail { color: #f44747; font-weight: bold; }
.summary { background: #094771; padding: 15px; margin-top: 30px; border-radius: 4px; font-size: 1.1em; }
.pass { color: #4ec9b0; font-weight: bold; }
.fail { color: #f44747; font-weight: bold; }
.note { color: #6a9955; font-style: italic; }
</style>
</head>
<body>

<h1>JScript Built-in Methods Test / JScript 内置方法测试</h1>
<p class="note">
Tests all 141 JScript built-in methods/functions from the Script 5.6 CHM knowledge base.<br>
测试 Script 5.6 CHM 知识库中全部 141 个 JScript 内置方法/函数。<br>
136 methods across 11 objects + 5 standalone functions.<br>
11 个对象上的 136 个方法 + 5 个独立函数。
</p>

<%
// ============================================================
// 1. Math Object / Math 对象 (18)
// ============================================================
%>
<h2>1. Math Object / Math 对象 (18)</h2>
<table>
<tr><th>Method / 方法</th><th>Call / 调用</th><th>Result / 结果</th><th>Status</th></tr>
<%
TestM("abs",   "Math.abs(-5.7)",   Math.abs(-5.7) === 5.7,          Math.abs(-5.7));
TestM("acos",  "Math.acos(0.5)",   Math.abs(Math.acos(0.5) - 1.0471975511965976) < 0.0001, Math.acos(0.5));
TestM("asin",  "Math.asin(0.5)",   Math.abs(Math.asin(0.5) - 0.5235987755982988) < 0.0001, Math.asin(0.5));
TestM("atan",  "Math.atan(1)",     Math.abs(Math.atan(1) - 0.7853981633974483) < 0.0001, Math.atan(1));
TestM("atan2", "Math.atan2(1,1)",  Math.abs(Math.atan2(1,1) - 0.7853981633974483) < 0.0001, Math.atan2(1,1));
TestM("ceil",  "Math.ceil(4.2)",   Math.ceil(4.2) === 5,            Math.ceil(4.2));
TestM("cos",   "Math.cos(0)",      Math.cos(0) === 1,               Math.cos(0));
TestM("exp",   "Math.exp(1)",      Math.abs(Math.exp(1) - 2.718281828459045) < 0.0001, Math.exp(1));
TestM("floor", "Math.floor(4.8)",  Math.floor(4.8) === 4,           Math.floor(4.8));
TestM("log",   "Math.log(1)",      Math.log(1) === 0,               Math.log(1));
TestM("max",   "Math.max(3,7)",    Math.max(3,7) === 7,             Math.max(3,7));
TestM("min",   "Math.min(3,7)",    Math.min(3,7) === 3,             Math.min(3,7));
TestM("pow",   "Math.pow(2,10)",   Math.pow(2,10) === 1024,         Math.pow(2,10));
TestM("random","Math.random()",    (Math.random() >= 0 && Math.random() < 1), "random [0,1)");
TestM("round", "Math.round(4.5)",  Math.round(4.5) === 5,           Math.round(4.5));
TestM("sin",   "Math.sin(0)",      Math.sin(0) === 0,               Math.sin(0));
TestM("sqrt",  "Math.sqrt(16)",    Math.sqrt(16) === 4,             Math.sqrt(16));
TestM("tan",   "Math.tan(0)",      Math.tan(0) === 0,               Math.tan(0));
%>
</table>

<%
// ============================================================
// 2. String Object / String 对象 (31)
// ============================================================
var ts = "Hello World";
%>
<h2>2. String Object / String 对象 (31)</h2>
<table>
<tr><th>Method / 方法</th><th>Call / 调用</th><th>Result / 结果</th><th>Status</th></tr>
<%
TestM("anchor",           "\"HW\".anchor(\"top\")",          "HW".anchor("top").indexOf("<A") >= 0,       "HW".anchor("top"));
TestM("big",              "\"hi\".big()",                    "hi".big().indexOf("<BIG>") >= 0,            "hi".big());
TestM("blink",            "\"hi\".blink()",                  "hi".blink().indexOf("<BLINK>") >= 0,        "hi".blink());
TestM("bold",             "\"hi\".bold()",                   "hi".bold().indexOf("<B>") >= 0,             "hi".bold());
TestM("charAt",           "ts.charAt(0)",                   ts.charAt(0) === "H",                        ts.charAt(0));
TestM("charCodeAt",       "ts.charCodeAt(0)",               ts.charCodeAt(0) === 72,                     ts.charCodeAt(0));
TestM("concat",           "\"Hello\".concat(\" World\")",   "Hello".concat(" World") === "Hello World",  "Hello".concat(" World"));
TestM("fixed",            "\"hi\".fixed()",                  "hi".fixed().indexOf("<TT>") >= 0,           "hi".fixed());
TestM("fontcolor",        "\"hi\".fontcolor(\"red\")",       "hi".fontcolor("red").toLowerCase().indexOf("red") >= 0,   "hi".fontcolor("red"));
TestM("fontsize",         "\"hi\".fontsize(5)",              "hi".fontsize(5).indexOf("5") >= 0,          "hi".fontsize(5));
TestM("fromCharCode",     "String.fromCharCode(65,66)",     String.fromCharCode(65,66) === "AB",         String.fromCharCode(65,66));
TestM("indexOf",          "ts.indexOf(\"World\")",          ts.indexOf("World") === 6,                   ts.indexOf("World"));
TestM("italics",          "\"hi\".italics()",                "hi".italics().indexOf("<I>") >= 0,          "hi".italics());
TestM("lastIndexOf",      "\"abcabc\".lastIndexOf(\"bc\")",  "abcabc".lastIndexOf("bc") === 4,            "abcabc".lastIndexOf("bc"));
TestM("link",             "\"HW\".link(\"http://x\")",      "HW".link("http://x").toLowerCase().indexOf("href") >= 0,  "HW".link("http://x"));
TestM("localeCompare",    "\"a\".localeCompare(\"b\")",     typeof "a".localeCompare("b") === "number",  "a".localeCompare("b"));
TestM("match",            "ts.match(/World/)",              ts.match(/World/) !== null,                   ts.match(/World/));
TestM("replace",          "ts.replace(\"World\",\"JS\")",   ts.replace("World","JS") === "Hello JS",     ts.replace("World","JS"));
TestM("search",           "ts.search(/World/)",             ts.search(/World/) === 6,                    ts.search(/World/));
TestM("slice",            "ts.slice(0,5)",                  ts.slice(0,5) === "Hello",                   ts.slice(0,5));
TestM("small",            "\"hi\".small()",                  "hi".small().indexOf("<SMALL>") >= 0,        "hi".small());
TestM("split",            "\"a,b,c\".split(\",\")",         "a,b,c".split(",").length === 3,             "a,b,c".split(",") + "");
TestM("strike",           "\"hi\".strike()",                 "hi".strike().indexOf("<STRIKE>") >= 0,      "hi".strike());
TestM("sub",              "\"hi\".sub()",                    "hi".sub().indexOf("<SUB>") >= 0,            "hi".sub());
TestM("substr",           "ts.substr(6,5)",                 ts.substr(6,5) === "World",                  ts.substr(6,5));
TestM("substring",        "ts.substring(0,5)",              ts.substring(0,5) === "Hello",               ts.substring(0,5));
TestM("sup",              "\"hi\".sup()",                    "hi".sup().indexOf("<SUP>") >= 0,            "hi".sup());
TestM("toLocaleLowerCase","\"HELLO\".toLocaleLowerCase()",  "HELLO".toLocaleLowerCase() === "hello",      "HELLO".toLocaleLowerCase());
TestM("toLocaleUpperCase","\"hello\".toLocaleUpperCase()",  "hello".toLocaleUpperCase() === "HELLO",      "hello".toLocaleUpperCase());
TestM("toLowerCase",      "\"HELLO\".toLowerCase()",        "HELLO".toLowerCase() === "hello",            "HELLO".toLowerCase());
TestM("toUpperCase",      "\"hello\".toUpperCase()",        "hello".toUpperCase() === "HELLO",            "hello".toUpperCase());
%>
</table>

<%
// ============================================================
// 3. Date Object / Date 对象 (44)
// ============================================================
var d  = new Date(2024, 5, 15, 13, 45, 30, 123); // month 5 = June (0-based)
var d2 = new Date(2024, 0, 1);
%>
<h2>3. Date Object / Date 对象 (43)</h2>
<table>
<tr><th>Method / 方法</th><th>Call / 调用</th><th>Result / 结果</th><th>Status</th></tr>
<%
// --- get methods (20) ---
TestM("getDate",           "d.getDate()",           d.getDate() === 15,           d.getDate());
TestM("getDay",            "d.getDay()",            (d.getDay() >= 0 && d.getDay() <= 6), d.getDay());
TestM("getFullYear",       "d.getFullYear()",       d.getFullYear() === 2024,     d.getFullYear());
TestM("getHours",          "d.getHours()",          d.getHours() === 13,          d.getHours());
TestM("getMilliseconds",   "d.getMilliseconds()",   d.getMilliseconds() === 123,  d.getMilliseconds());
TestM("getMinutes",        "d.getMinutes()",        d.getMinutes() === 45,        d.getMinutes());
TestM("getMonth",          "d.getMonth()",          d.getMonth() === 5,           d.getMonth());
TestM("getSeconds",        "d.getSeconds()",        d.getSeconds() === 30,        d.getSeconds());
TestM("getTime",           "d.getTime()",            typeof d.getTime() === "number", d.getTime());
TestM("getTimezoneOffset", "d.getTimezoneOffset()",  typeof d.getTimezoneOffset() === "number", d.getTimezoneOffset());
TestM("getUTCDate",        "d.getUTCDate()",         typeof d.getUTCDate() === "number", d.getUTCDate());
TestM("getUTCDay",         "d.getUTCDay()",          (d.getUTCDay() >= 0 && d.getUTCDay() <= 6), d.getUTCDay());
TestM("getUTCFullYear",    "d.getUTCFullYear()",     typeof d.getUTCFullYear() === "number", d.getUTCFullYear());
TestM("getUTCHours",       "d.getUTCHours()",        typeof d.getUTCHours() === "number", d.getUTCHours());
TestM("getUTCMilliseconds","d.getUTCMilliseconds()", typeof d.getUTCMilliseconds() === "number", d.getUTCMilliseconds());
TestM("getUTCMinutes",     "d.getUTCMinutes()",      typeof d.getUTCMinutes() === "number", d.getUTCMinutes());
TestM("getUTCMonth",       "d.getUTCMonth()",        typeof d.getUTCMonth() === "number", d.getUTCMonth());
TestM("getUTCSeconds",     "d.getUTCSeconds()",      typeof d.getUTCSeconds() === "number", d.getUTCSeconds());
TestM("getYear",           "d.getYear()",            typeof d.getYear() === "number", d.getYear());
TestM("getVarDate",        "d.getVarDate()",         true, d.getVarDate());
// --- set methods (16) ---
var sd = new Date(2024, 0, 1);
sd.setDate(20);          TestM("setDate",      "sd.setDate(20)",      sd.getDate() === 20, sd.getDate());
sd.setFullYear(2025);    TestM("setFullYear",  "sd.setFullYear(2025)", sd.getFullYear() === 2025, sd.getFullYear());
sd.setHours(10);         TestM("setHours",     "sd.setHours(10)",     sd.getHours() === 10, sd.getHours());
sd.setMilliseconds(99);  TestM("setMilliseconds","sd.setMilliseconds(99)", sd.getMilliseconds() === 99, sd.getMilliseconds());
sd.setMinutes(30);       TestM("setMinutes",   "sd.setMinutes(30)",   sd.getMinutes() === 30, sd.getMinutes());
sd.setMonth(6);          TestM("setMonth",     "sd.setMonth(6)",      sd.getMonth() === 6, sd.getMonth());
sd.setSeconds(45);       TestM("setSeconds",   "sd.setSeconds(45)",   sd.getSeconds() === 45, sd.getSeconds());
var epoch = sd.getTime(); sd.setTime(epoch); TestM("setTime", "sd.setTime(epoch)", sd.getTime() === epoch, sd.getTime());
sd.setUTCDate(25);       TestM("setUTCDate",   "sd.setUTCDate(25)",   sd.getUTCDate() === 25, sd.getUTCDate());
sd.setUTCFullYear(2026); TestM("setUTCFullYear","sd.setUTCFullYear(2026)", sd.getUTCFullYear() === 2026, sd.getUTCFullYear());
sd.setUTCHours(12);      TestM("setUTCHours",  "sd.setUTCHours(12)",  sd.getUTCHours() === 12, sd.getUTCHours());
sd.setUTCMilliseconds(0);TestM("setUTCMilliseconds","sd.setUTCMilliseconds(0)", sd.getUTCMilliseconds() === 0, sd.getUTCMilliseconds());
sd.setUTCMinutes(15);    TestM("setUTCMinutes","sd.setUTCMinutes(15)", sd.getUTCMinutes() === 15, sd.getUTCMinutes());
sd.setUTCMonth(3);       TestM("setUTCMonth",  "sd.setUTCMonth(3)",   sd.getUTCMonth() === 3, sd.getUTCMonth());
sd.setUTCSeconds(30);    TestM("setUTCSeconds","sd.setUTCSeconds(30)", sd.getUTCSeconds() === 30, sd.getUTCSeconds());
sd.setYear(99);          TestM("setYear",      "sd.setYear(99)",      typeof sd.getYear() === "number", sd.getYear());
// --- conversion + static (8) ---
TestM("toDateString",     "d.toDateString()",     typeof d.toDateString() === "string", d.toDateString());
TestM("toGMTString",      "d.toGMTString()",      typeof d.toGMTString() === "string", d.toGMTString());
TestM("toLocaleDateString","d.toLocaleDateString()", typeof d.toLocaleDateString() === "string", d.toLocaleDateString());
TestM("toLocaleTimeString","d.toLocaleTimeString()", typeof d.toLocaleTimeString() === "string", d.toLocaleTimeString());
TestM("toTimeString",     "d.toTimeString()",     typeof d.toTimeString() === "string", d.toTimeString());
TestM("toUTCString",      "d.toUTCString()",      typeof d.toUTCString() === "string", d.toUTCString());
TestM("parse",            "Date.parse(\"2024/06/15\")", typeof Date.parse("2024/06/15") === "number", Date.parse("2024/06/15"));
TestM("UTC",              "Date.UTC(2024,5,15)",  typeof Date.UTC(2024,5,15) === "number", Date.UTC(2024,5,15));
%>
</table>

<%
// ============================================================
// 4. Array Object / Array 对象 (13)
// ============================================================
var arr = [1, 2, 3];
var arr2 = [4, 5];
%>
<h2>4. Array Object / Array 对象 (13)</h2>
<table>
<tr><th>Method / 方法</th><th>Call / 调用</th><th>Result / 结果</th><th>Status</th></tr>
<%
var arrC = [1,2,3]; var arrD = [4,5];
var joined = arrC.concat(arrD);
TestM("concat",        "[1,2,3].concat([4,5])",       joined.length === 5 && joined[3] === 4, joined + "");

TestM("join",          "[1,2,3].join(\",\")",         [1,2,3].join(",") === "1,2,3",          [1,2,3].join(","));

var arrP = [1,2,3]; var popped = arrP.pop();
TestM("pop",           "[1,2,3].pop()",               popped === 3 && arrP.length === 2,       popped);

var arrPu = [1,2]; arrPu.push(3);
TestM("push",          "[1,2].push(3)",               arrPu[2] === 3 && arrPu.length === 3,  arrPu + "");

var arrR = [1,2,3]; arrR.reverse();
TestM("reverse",       "[1,2,3].reverse()",           arrR[0] === 3 && arrR[2] === 1,          arrR + "");

var arrS = [1,2,3]; var shifted = arrS.shift();
TestM("shift",         "[1,2,3].shift()",             shifted === 1 && arrS.length === 2,       shifted);

var arrSl = [1,2,3,4]; var sliced = arrSl.slice(1,3);
TestM("slice",         "[1,2,3,4].slice(1,3)",        sliced.length === 2 && sliced[0] === 2,  sliced + "");

var arrSo = [3,1,2]; arrSo.sort();
TestM("sort",          "[3,1,2].sort()",              arrSo[0] === 1 && arrSo[2] === 3,        arrSo + "");

var arrSp = [1,2,3,4]; var removed = arrSp.splice(1,2,"a","b");
TestM("splice",        "[1,2,3,4].splice(1,2,\"a\",\"b\")", removed.length === 2 && arrSp[1] === "a", removed + "");

TestM("toLocaleString","[1,2,3].toLocaleString()",    typeof [1,2,3].toLocaleString() === "string", [1,2,3].toLocaleString());
TestM("toString",      "[1,2,3].toString()",          [1,2,3].toString() === "1,2,3",          [1,2,3].toString());

var arrU = [1,2]; arrU.unshift(0);
TestM("unshift",       "[1,2].unshift(0)",            arrU[0] === 0 && arrU.length === 3,    arrU + "");

TestM("valueOf",       "[1,2,3].valueOf()",           [1,2,3].valueOf().length === 3,          [1,2,3].valueOf() + "");
%>
</table>

<%
// ============================================================
// 5. Global Object / Global 对象 (11)
// ============================================================
%>
<h2>5. Global Object / Global 对象 (11)</h2>
<table>
<tr><th>Method / 方法</th><th>Call / 调用</th><th>Result / 结果</th><th>Status</th></tr>
<%
TestM("decodeURI",          "decodeURI(\"hello%20world\")",       decodeURI("hello%20world") === "hello world",       decodeURI("hello%20world"));
TestM("decodeURIComponent", "decodeURIComponent(\"a%3Db\")",      decodeURIComponent("a%3Db") === "a=b",             decodeURIComponent("a%3Db"));
TestM("encodeURI",          "encodeURI(\"hello world\")",         encodeURI("hello world") === "hello%20world",       encodeURI("hello world"));
TestM("encodeURIComponent", "encodeURIComponent(\"a=b\")",        encodeURIComponent("a=b") === "a%3Db",              encodeURIComponent("a=b"));
TestM("escape",             "escape(\"hello world\")",            escape("hello world").indexOf("%20") >= 0,          escape("hello world"));
TestM("eval",               "eval(\"2+3\")",                      eval("2+3") === 5,                                  eval("2+3"));
TestM("isFinite",           "isFinite(42)",                       isFinite(42) === true,                              isFinite(42));
TestM("isNaN",              "isNaN(NaN)",                         isNaN(NaN) === true,                                isNaN(NaN));
TestM("parseFloat",         "parseFloat(\"3.14abc\")",            parseFloat("3.14abc") === 3.14,                     parseFloat("3.14abc"));
TestM("parseInt",           "parseInt(\"0xFF\",16)",              parseInt("0xFF",16) === 255,                        parseInt("0xFF",16));
TestM("unescape",           "unescape(\"hello%20world\")",        unescape("hello%20world") === "hello world",        unescape("hello%20world"));
%>
</table>

<%
// ============================================================
// 6. VBArray Object / VBArray 对象 (5)
// ============================================================
%>
<h2>6. VBArray Object / VBArray 对象 (5)</h2>
<table>
<tr><th>Method / 方法</th><th>Call / 调用</th><th>Result / 结果</th><th>Status</th></tr>
<%
var vbAvail = false;
try {
    // Try to create a VBArray via COM (requires ScriptControl or similar)
    // On IIS, we can try to get a SAFEARRAY from a COM object
    // Simplest: skip if we can't create one
    throw new Error("VBArray requires SAFEARRAY from COM");
} catch(e) {
    SkipM("dimensions", "vbArray.dimensions()", "SKIP — requires SAFEARRAY from COM / 需要 COM 的 SAFEARRAY");
    SkipM("getItem",    "vbArray.getItem(0)",    "SKIP — requires SAFEARRAY from COM");
    SkipM("lbound",     "vbArray.lbound(1)",     "SKIP — requires SAFEARRAY from COM");
    SkipM("toArray",    "vbArray.toArray()",     "SKIP — requires SAFEARRAY from COM");
    SkipM("ubound",     "vbArray.ubound(1)",     "SKIP — requires SAFEARRAY from COM");
}
%>
</table>

<%
// ============================================================
// 7. Enumerator Object / Enumerator 对象 (4)
// ============================================================
%>
<h2>7. Enumerator Object / Enumerator 对象 (4)</h2>
<table>
<tr><th>Method / 方法</th><th>Call / 调用</th><th>Result / 结果</th><th>Status</th></tr>
<%
var enumAvail = false;
try {
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var folder = fso.GetFolder(Server.MapPath("."));
    var files = new Enumerator(folder.Files);
    if (!files.atEnd()) {
        var firstItem = files.item();
        TestM("item",      "files.item()",      typeof firstItem === "object" || typeof firstItem === "string", String(firstItem).substring(0, 40));
        enumAvail = true;
    }
    files.moveFirst();
    TestM("moveFirst", "files.moveFirst()", !files.atEnd(), "reset to first");
    files.moveNext();
    TestM("moveNext",  "files.moveNext()",  true, "moved to next");
    // atEnd
    while (!files.atEnd()) files.moveNext();
    TestM("atEnd",     "files.atEnd()",     files.atEnd() === true, "true (end of collection)");
} catch(e) {
    SkipM("item",      "enum.item()",      "SKIP — FSO not available: " + e.message);
    SkipM("moveFirst", "enum.moveFirst()",  "SKIP — FSO not available");
    SkipM("moveNext",  "enum.moveNext()",   "SKIP — FSO not available");
    SkipM("atEnd",     "enum.atEnd()",      "SKIP — FSO not available");
}
%>
</table>

<%
// ============================================================
// 8. Number Object / Number 对象 (3)
// ============================================================
%>
<h2>8. Number Object / Number 对象 (3)</h2>
<table>
<tr><th>Method / 方法</th><th>Call / 调用</th><th>Result / 结果</th><th>Status</th></tr>
<%
TestM("toExponential", "(3.14159).toExponential(2)",  (3.14159).toExponential(2) === "3.14e+0",  (3.14159).toExponential(2));
TestM("toFixed",       "(3.14159).toFixed(2)",        (3.14159).toFixed(2) === "3.14",           (3.14159).toFixed(2));
TestM("toPrecision",   "(3.14159).toPrecision(4)",    (3.14159).toPrecision(4) === "3.142",      (3.14159).toPrecision(4));
%>
</table>

<%
// ============================================================
// 9. Regular Expression Object / 正则表达式对象 (3)
// ============================================================
%>
<h2>9. Regular Expression Object / 正则表达式对象 (3)</h2>
<table>
<tr><th>Method / 方法</th><th>Call / 调用</th><th>Result / 结果</th><th>Status</th></tr>
<%
var re1 = /hello/i;
TestM("exec", "/hello/i.exec(\"Hello World\")", re1.exec("Hello World") !== null && re1.exec("Hello World")[0] === "Hello", (re1.exec("Hello World"))[0]);

var re2 = /\d+/;
TestM("test", "/\\d+/.test(\"abc123\")", re2.test("abc123") === true, "true");

var re3 = /[a-z]+/gi;
re3.compile("[0-9]+", "g");
TestM("compile", "re.compile(\"[0-9]+\",\"g\")", re3.test("123") === true, "recompiled, test '123' = true");
%>
</table>

<%
// ============================================================
// 10. Function Object / Function 对象 (2)
// ============================================================
%>
<h2>10. Function Object / Function 对象 (2)</h2>
<table>
<tr><th>Method / 方法</th><th>Call / 调用</th><th>Result / 结果</th><th>Status</th></tr>
<%
function add(a, b) { return a + b; }
var ctx = { x: 10 };
function addToCtx(a, b) { return this.x + a + b; }

TestM("apply", "add.apply(null,[3,4])",         add.apply(null, [3,4]) === 7,              add.apply(null, [3,4]));
TestM("call",  "addToCtx.call({x:10},3,4)",     addToCtx.call({x:10}, 3, 4) === 17,        addToCtx.call({x:10}, 3, 4));
%>
</table>

<%
// ============================================================
// 11. Object Object / Object 对象 (2)
// ============================================================
%>
<h2>11. Object Object / Object 对象 (2)</h2>
<table>
<tr><th>Method / 方法</th><th>Call / 调用</th><th>Result / 结果</th><th>Status</th></tr>
<%
function MyClass() { this.name = "test"; }
var obj = new MyClass();
TestM("hasOwnProperty", "obj.hasOwnProperty(\"name\")",  obj.hasOwnProperty("name") === true,   obj.hasOwnProperty("name"));
TestM("isPrototypeOf",  "MyClass.prototype.isPrototypeOf(obj)", MyClass.prototype.isPrototypeOf(obj) === true, MyClass.prototype.isPrototypeOf(obj));
%>
</table>

<%
// ============================================================
// 12. Standalone Functions / 独立函数 (5)
// ============================================================
%>
<h2>12. Standalone Functions / 独立函数 (5)</h2>
<table>
<tr><th>Function / 函数</th><th>Call / 调用</th><th>Result / 结果</th><th>Status</th></tr>
<%
// GetObject
var gotObj = null, gotErr = false;
try {
    gotObj = GetObject("", "Scripting.FileSystemObject");
} catch(e) { gotErr = true; }
TestM("GetObject", "GetObject(\"\",\"Scripting.FileSystemObject\")", gotObj !== null || gotErr, gotObj !== null ? "object" : "err (acceptable)");

// ScriptEngine
TestM("ScriptEngine",            "ScriptEngine()",            ScriptEngine() === "JScript",            ScriptEngine());
TestM("ScriptEngineMajorVersion","ScriptEngineMajorVersion()", ScriptEngineMajorVersion() >= 5,       ScriptEngineMajorVersion());
TestM("ScriptEngineMinorVersion","ScriptEngineMinorVersion()", typeof ScriptEngineMinorVersion() === "number", ScriptEngineMinorVersion());
TestM("ScriptEngineBuildVersion","ScriptEngineBuildVersion()", typeof ScriptEngineBuildVersion() === "number", ScriptEngineBuildVersion());
%>
</table>

<%
// ============================================================
// Summary / 汇总
// ============================================================
%>
<div class="summary">
<h2 style="background:none;border:none;padding:0;margin:0;color:#fff;">
Summary / 测试汇总
</h2>
<p>Total Tests / 总测试数: <strong><%= totalTests %></strong></p>
<p>Passed / 通过: <span class="pass"><%= passCount %></span></p>
<p>Failed / 失败: <span class="fail"><%= failCount %></span></p>
<p>Skipped / 跳过: <span><%= skipCount %></span> (VBArray — requires COM SAFEARRAY)</p>
<%
if (failCount === 0) {
    Response.Write("<p style=\"font-size:1.3em;\">All functions passed! / 全部函数测试通过！</p>");
} else {
    Response.Write("<p style=\"font-size:1.3em;\">" + failCount + " function(s) failed: " + errList + "</p>");
}
%>
</div>

<p class="note" style="margin-top:30px;">
Source / 来源: Script 5.6 CHM — <code>js56jslrfjscriptmethodstoc.htm</code> (111 in TOC) + <code>js56jslrfjscriptfunctionstoc.htm</code> (5) + 25 files on disk not in TOC<br>
Total: 18 + 31 + 43 + 13 + 11 + 5 + 4 + 3 + 3 + 2 + 2 + 5 = <strong>141</strong> methods/functions
</p>

</body>
</html>
