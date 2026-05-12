<%
    Dim tempoInicio, tempoFim, totalSegundos
    tempoInicio = Timer

    Sub DoSomething()
    Dim i
    For i = 1 To 1000000
    Next
    End Sub

    DoSomething()

    ' Captura o tempo final
    tempoFim = Timer

    ' Calcula a diferença
    totalSegundos = tempoFim - tempoInicio
    response.write "Tempo gasto para executar o loop: " & totalSegundos & " segundos."
%>


<script runat="server" language="JScript">

    // Captura o tempo inicial (em milissegundos)
    var tempoInicio = new Date().getTime();

    // O Loop solicitado
    for (var i = 1; i <= 1000000; i++) {
        // Apenas iterando.
    }

    // Captura o tempo final
    var tempoFim = new Date().getTime();

    // Calcula a diferença e converte para segundos
    // Dividimos por 1000 porque o JS trabalha com milissegundos
    var totalSegundos = (tempoFim - tempoInicio) / 1000;

    Response.Write("Tempo gasto para executar o loop: " + totalSegundos + " segundos.");


</script>