<%
Dim startTime, endTime, executionTime
Dim i, j, isPrime, limit, primeCount
Dim primes()

limit = 50000
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
        ' Realocação de memória e cópia de dados a cada novo número primo encontrado.
        ReDim Preserve primes(primeCount)
        primes(primeCount) = i
        primeCount = primeCount + 1
    End If
Next

endTime = Timer()
' A função Timer retorna segundos desde a meia-noite. Multiplicar por 1000 para ms.
executionTime = (endTime - startTime) * 1000

Console.log "VBScript Execution Time (ReDim Preserve): " & Round(executionTime, 2) & " ms"
Console.log "Primes found: " & primeCount
%>