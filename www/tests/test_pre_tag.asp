<%@ Language=VBScript %>
<%
Option Explicit

Function pre(value)
    pre = Replace(value, vbTab, " ")
End Function

%>
<!DOCTYPE html>
<html>
<head>
    <title>Test with pre Tag</title>
</head>
<body>
<h1>Test with &lt;pre&gt; Tag</h1>

<p>Test 1: Normal output</p>
<p><%= pre("test") %></p>

<p>Test 2: Inside pre tag</p>
<pre><%= pre("test2") %></pre>

<p>Test 3: Inside pre tag with attributes</p>
<pre class="code"><%= pre("test3") %></pre>

</body>
</html>
