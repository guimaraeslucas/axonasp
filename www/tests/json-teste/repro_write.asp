<%
Response.Buffer = True
%>
<!--#include file="jsonObject.class.asp" -->
<%
Dim jsonObj, outputObj
Set jsonObj = New JSONobject
Set outputObj = jsonObj.parse("{""a"":1}")
%>
<pre><%= outputObj.Write %></pre>
