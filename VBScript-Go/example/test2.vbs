' Estruturas de controle complexas

Dim i, j, total
total = 0

' Loop For aninhado
For i = 1 To 5
    For j = 1 To 3
        total = total + (i * j)
    Next
Next

Response.Write "Total: " & total & "<br>"

' Select Case
Dim nota
nota = 85

Select Case nota
    Case 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100
        Response.Write "Nota A"
    Case 80, 81, 82, 83, 84, 85, 86, 87, 88, 89
        Response.Write "Nota B"
    Case 70, 71, 72, 73, 74, 75, 76, 77, 78, 79
        Response.Write "Nota C"
    Case Else
        Response.Write "Nota D ou F"
End Select

' Do While
Dim contador
contador = 0

Do While contador < 10
    contador = contador + 2
    If contador = 6 Then
        Response.Write "Pulando 6<br>"
    Else
        Response.Write "Contador: " & contador & "<br>"
    End If
Loop
