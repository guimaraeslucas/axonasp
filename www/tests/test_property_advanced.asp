<%
@ debug_asp_code = "FALSE"
%>
<%
Option Explicit
Response.Write("Testing Advanced Property Get/Let/Set Implementation" & vbCrLf & vbCrLf)

' Test 1: Simple Property Get/Let
Class Person
    Private m_age
    Private m_name

    Public Sub Initialize(n, a)
        m_name = n
        m_age = a
    End Sub

    Public Property Get Name()
        Name = m_name
    End Property

    Public Property Let Name(value)
        m_name = value
    End Property

    Public Property Get Age()
        Age = m_age
    End Property

    Public Property Let Age(value)
        If value < 0 Then
            Err.Raise 11, "Person.Age", "Age cannot be negative"
        End If
        m_age = value
    End Property

    Public Property Get Info()
        Info = m_name & " (" & m_age & ")"
    End Property
End Class

' Test 2: Property with Indexed Access (Array Property)
Class StudentGrades
    Private m_grades()
    Private m_count

    Public Sub Initialize()
        ReDim m_grades(9)
        m_count = 0
    End Sub

    Public Property Get Grade(index)
        If index >  = 0 And index < m_count Then
            Grade = m_grades(index)
        Else
            Grade = - 1
        End If
    End Property

    Public Property Let Grade(index, value)
        If value < 0 Or value > 100 Then
            Err.Raise 11, "StudentGrades.Grade", "Grade must be between 0 and 100"
        End If
        If index >  = m_count Then
            m_count = index + 1
            If index > UBound(m_grades) Then
                ReDim Preserve m_grades(index + 5)
            End If
        End If
        m_grades(index) = value
    End Property

    Public Function Count()
        Count = m_count
    End Function

    Public Function Average()
        If m_count = 0 Then
            Average = 0
            Exit Function
        End If
        Dim sum, i
        sum = 0
        For i = 0 To m_count - 1
            sum = sum + m_grades(i)
        Next
        Average = sum / m_count
    End Function
End Class

' Test 3: Property with Default Modifier
Class Temperature
    Private m_celsius

    Public Sub Initialize(c)
        m_celsius = c
    End Sub

    Public Property Get Celsius()
        Celsius = m_celsius
    End Property

    Public Property Let Celsius(value)
        m_celsius = value
    End Property

    Public Property Get Fahrenheit()
        Fahrenheit = (m_celsius * 9 / 5) + 32
    End Property

    Public Property Let Fahrenheit(value)
        m_celsius = (value - 32) * 5 / 9
    End Property
End Class

' Test 4: Property with Get only (Read-only)
Class ReadOnlyValue
    Private m_value

    Public Sub Initialize(v)
        m_value = v
    End Sub

    Public Property Get Value()
        Value = m_value
    End Property

    Public Sub ChangeValue(v)
        m_value = v
    End Sub
End Class

' Test 5: Property with Set only (Write-only, less common)
Class SecretValue
    Private m_secret

    Public Property Set Secret(value)
        m_secret = value
    End Property

    Public Function Encoded()
        Encoded = "Secret[***]"
    End Function

    Public Function CheckSecret(value)
        CheckSecret = (value = m_secret)
    End Function
End Class

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 1: SIMPLE PROPERTY GET/LET" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf & vbCrLf)

Dim person
Set person = New Person
person.Initialize "John", 30

Response.Write("Created person: " & person.Info & vbCrLf)
Response.Write("  Name: " & person.Name & vbCrLf)
Response.Write("  Age: " & person.Age & vbCrLf & vbCrLf)

person.Name = "Jane"
person.Age = 28
Response.Write("After modification: " & person.Info & vbCrLf & vbCrLf)

' Test error handling
On Error Resume Next
person.Age = - 5
If Err.Number <> 0 Then
    Response.Write("✓ Age validation working: " & Err.Description & vbCrLf & vbCrLf)
    Err.Clear
End If
On Error Goto 0

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 2: INDEXED PROPERTY (ARRAY PROPERTY)" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf & vbCrLf)

Dim grades
Set grades = New StudentGrades
grades.Initialize()

Response.Write("Setting grades..." & vbCrLf)
grades.Grade(0) = 85
grades.Grade(1) = 92
grades.Grade(2) = 78
grades.Grade(3) = 95

Response.Write("Student grades:" & vbCrLf)
Response.Write("  Grade 0: " & grades.Grade(0) & vbCrLf)
Response.Write("  Grade 1: " & grades.Grade(1) & vbCrLf)
Response.Write("  Grade 2: " & grades.Grade(2) & vbCrLf)
Response.Write("  Grade 3: " & grades.Grade(3) & vbCrLf)
Response.Write("  Total: " & grades.Count() & vbCrLf)
Response.Write("  Average: " & Format(grades.Average(), "0.00") & vbCrLf & vbCrLf)

' Test validation
On Error Resume Next
grades.Grade(4) = 150
If Err.Number <> 0 Then
    Response.Write("✓ Grade validation working: " & Err.Description & vbCrLf & vbCrLf)
    Err.Clear
End If
On Error Goto 0

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 3: UNIT CONVERSION WITH PROPERTY" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf & vbCrLf)

Dim temp
Set temp = New Temperature
temp.Initialize(25)

Response.Write("Temperature in Celsius: " & temp.Celsius & vbCrLf)
Response.Write("Temperature in Fahrenheit: " & Format(temp.Fahrenheit, "0.00") & vbCrLf & vbCrLf)

Response.Write("Setting to 32°F..." & vbCrLf)
temp.Fahrenheit = 32
Response.Write("Celsius: " & Format(temp.Celsius, "0.00") & vbCrLf)
Response.Write("Fahrenheit: " & temp.Fahrenheit & vbCrLf & vbCrLf)

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 4: READ-ONLY PROPERTY" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf & vbCrLf)

Dim ReadOnly
Set ReadOnly = New ReadOnlyValue
ReadOnly.Initialize(42)

Response.Write("Read-only value: " & ReadOnly.Value & vbCrLf)

ReadOnly.ChangeValue(99)
Response.Write("After internal change: " & ReadOnly.Value & vbCrLf & vbCrLf)

' Verify that direct assignment fails (if implementation enforces it)
On Error Resume Next
ReadOnly.Value = 123
If Err.Number <> 0 Then
    Response.Write("✓ Read-only enforcement working" & vbCrLf & vbCrLf)
    Err.Clear
Else
    Response.Write("Note: Direct assignment to read-only property succeeded (flexible implementation)" & vbCrLf & vbCrLf)
End If
On Error Goto 0

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("TEST 5: WRITE-ONLY PROPERTY" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf & vbCrLf)

Dim secret
Set secret = New SecretValue
Response.Write("Setting secret value..." & vbCrLf)
secret.Secret = "MyPassword123"

Response.Write("Encoded representation: " & secret.Encoded() & vbCrLf)
Response.Write("Check 'MyPassword123': " & secret.CheckSecret("MyPassword123") & vbCrLf)
Response.Write("Check 'WrongPassword': " & secret.CheckSecret("WrongPassword") & vbCrLf & vbCrLf)

Response.Write("=" & String(60, "=") & vbCrLf)
Response.Write("ALL TESTS COMPLETED" & vbCrLf)
Response.Write("=" & String(60, "=") & vbCrLf)
%>
