<p><strong><%
'This is a sample module that can be added to your QuickerSite. 
'You can have very simple scripts like response.write "Hello World", 
'but you can also develop a full blown application, link to a database, 
'a 3rd party application.
'For now, let's just do a simple trick:

Dim dHour
dHour = Hour(Now)

If dHour < 12 Then
	Response.Write "Good morning!"
ElseIf dHour < 17 Then
	Response.Write "Good afternoon!"
Else
	Response.Write "Good evening!"
End If

%></strong></p>