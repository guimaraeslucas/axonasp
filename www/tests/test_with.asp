<%@ Language=VBScript %>
<%
Response.Write "WITH_START<br>"

Dim dict
Set dict = CreateObject("Scripting.Dictionary")
With dict
    .Add "name", "axon"
    .Add "lang", "vb"
    Response.Write "DICT:" & .Item("name") & "|" & .Count & "<br>"
End With

Class Person
    Public Name
    Public Age
    Public Function Greeting()
        Greeting = Name & "-" & Age
    End Function
    Public Sub Birthday(years)
        Age = Age + years
    End Sub
End Class

Dim p
Set p = New Person
p.Name = "Ada"
p.Age = 41

With p
    .Birthday 1
    .Name = "Ada Lovelace"
    Response.Write "PERSON:" & .Greeting() & "|" & .Name & "|" & .Age & "<br>"
End With

Response.Write "WITH_END"
%>
