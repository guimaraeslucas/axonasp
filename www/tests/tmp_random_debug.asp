<%
Class cls_asplite_randomizer
    Private Sub Class_Initialize()
        Randomize()
    End Sub

    Public Function randomText(nmbrChars)
        Dim i
        For i = 1 To nmbrChars
            randomText = randomText & Chr(Int((26) * Rnd + 97))
        Next
    End Function

    Public Function randomNumber(startnr, stopnr)
        randomNumber = (stopnr - startnr + 1) * Rnd + startnr
    End Function
End Class

Dim r, code, i
Set r = New cls_asplite_randomizer
For i = 1 To 10
    code = Int((26) * Rnd + 97)
    Response.Write code & ":" & Chr(code) & "|"
Next
Response.Write "#" & r.randomText(20)
%>
