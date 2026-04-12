<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Classes Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .box { border-left: 4px solid #667eea; padding: 15px; margin-bottom: 15px; background: #f9f9f9; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Classes & OOP Test</h1>
        <div class="intro">
            <p>Tests class definition, properties, methods, constructors and inheritance patterns.</p>
        </div>
        <div class="box">

<%
Response.Write "<h3>Testing Classes</h3>"

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

' Test Scope Shadowing
Class ScopeTest
    Public X
    Public Sub Test(arg)
        X = 10
        Response.Write "Member X: " & X & "<br>" 
        Response.Write "Arg: " & arg & "<br>"
    End Sub
End Class

Dim s
Set s = New ScopeTest
s.Test "argument"
Response.Write "ScopeTest X: " & s.X & "<br>"

%>
        </div>
    </div>
</body>
</html>
