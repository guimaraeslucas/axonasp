Option Explicit

' Exemplo de funções e variáveis
Dim nome, idade, salario

nome = "João Silva"
idade = 30
salario = 5000.50

Function CalcularBonus(sal, percentual)
    CalcularBonus = sal * (percentual / 100)
End Function

Sub ExibirDados()
    Response.Write "Nome: " & nome & "<br>"
    Response.Write "Idade: " & idade & "<br>"
    Response.Write "Salário: " & salario & "<br>"
    
    Dim bonus
    bonus = CalcularBonus(salario, 10)
    Response.Write "Bônus (10%): " & bonus & "<br>"
End Sub

ExibirDados()
