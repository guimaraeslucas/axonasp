<%
Dim startTime, endTime, executionTime
Dim i, j, isPrime, limit, primeCount
Dim primesDict

limit = 50000
' Uso de Dictionary para evitar custo computacional de ReDim Preserve O(n)
Set primesDict = CreateObject("Scripting.Dictionary")
primeCount = 0

startTime = Timer()

For i = 2 To limit
    isPrime = True
    For j = 2 To Int(Sqr(i))
        If i Mod j = 0 Then
            isPrime = False
            Exit For
        End If
    Next
    
    If isPrime Then
        primesDict.Add primeCount, i
        primeCount = primeCount + 1
    End If
Next

endTime = Timer()
' A função Timer retorna segundos desde a meia-noite. Multiplicar por 1000 para ms.
executionTime = (endTime - startTime) * 1000

Console.log "VBScript Execution Time: " & Round(executionTime, 2) & " ms"
Console.log "Primes found: " & primesDict.Count

Set primesDict = Nothing
%>