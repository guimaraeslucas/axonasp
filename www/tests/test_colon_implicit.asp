<%@ Language=VBScript %>
<%
' Test Colon syntax and Implicit Set
Dim obj : Set obj = Server.CreateObject("Scripting.Dictionary")
obj.Add "Key", "Value"
Response.Write "Dictionary Item: " & obj("Key") & "<br>"

' Test Implicit Set (Abbreviated)
Dim obj2
obj2 = Server.CreateObject("Scripting.Dictionary") ' Should work without Set
obj2.Add "Key2", "Value2"
Response.Write "Implicit Set Item: " & obj2("Key2") & "<br>"

' Test Colon with Implicit Set
Dim obj3 : obj3 = Server.CreateObject("Scripting.Dictionary")
obj3.Add "Key3", "Value3"
Response.Write "Colon + Implicit Set Item: " & obj3("Key3") & "<br>"
%>
