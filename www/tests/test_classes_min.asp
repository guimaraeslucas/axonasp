<%@ Language=VBScript %>
<%
Response.Write "Starting min classes test...<br>"

Class Person
    Public Name
    Private age_
    
    Private Sub Class_Initialize()
        age_ = 0
        Name = "Unnamed"
    End Sub

    Public Property Get Age
        Age = age_
    End Property

    Public Property Let Age(v)
        if v >= 0 then age_ = v
    End Property
    
    Public Function Describe()
        Describe = Name & " is " & age_ & " years old."
    End Function
    
    Public Sub Birthday()
        age_ = age_ + 1
    End Sub
    
    Public Function SelfRef()
        SelfRef = Me.Name
    End Function
End Class

Dim p
Set p = New Person
Response.Write "Initial: " & p.Describe() & "<br>"

p.Name = "Alice"
p.Age = 30
Response.Write "Updated: " & p.Describe() & "<br>"

p.Birthday
Response.Write "After Birthday: " & p.Age & "<br>"
Response.Write "Self Ref: " & p.SelfRef() & "<br>"
%>
