<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' Test Class Dim Statements
' This test verifies that Dim statements inside classes work correctly

Response.Write "=== Testing Class Dim Statements ===" & vbCrLf

' Test 1: Basic Dim statements in class
Class TestClass1
    Dim name
    Dim age
    Dim scores(10)
    
    Public Sub setName(v)
        name = v
    End Sub
    
    Public Function getName()
        getName = name
    End Function
    
    Public Sub setAge(v)
        age = v
    End Sub
    
    Public Function getAge()
        getAge = age
    End Function
End Class

Dim obj1
Set obj1 = New TestClass1
obj1.setName "John"
obj1.setAge 25
Response.Write "Test1-Basic: name=" & obj1.getName() & " age=" & obj1.getAge() & vbCrLf

' Test 2: Conexao class simulation (the actual bug)
Class ConexaoSim
    Dim stringConexao
    Dim servidor
    Dim bancoDados
    Dim usuario
    Dim senha
    Dim sgbd
    Dim objConexao
    
    Public sub setSgbd(vSgbd)
        sgbd = vSgbd
    End sub
    
    Public sub setUsuario(vUsuario)
        usuario = vUsuario
    End sub
    
    Public sub setSenha(vSenha)
        senha = vSenha
    End sub
    
    Public sub setServidor(vServidor)
        servidor = vServidor
    End sub
    
    Public sub setBancoDados(vBancoDados)
        bancoDados = vBancoDados
    End sub
    
    Public Sub setStringConexao()
        Select Case sgbd
            Case "oracle"
                stringConexao = "oracle_connection_string_test"
            Case "sqlserver"
                stringConexao = "sqlserver_connection_string_test"
        End Select
    End Sub
    
    Public Function getStringConexao()
        getStringConexao = stringConexao
    End Function
End Class

Dim conecta
Set conecta = New ConexaoSim

' Set variables like the real code does
conecta.setSgbd("oracle")
conecta.setUsuario("test_user")
conecta.setSenha("test_pass")
conecta.setServidor("test_server")
conecta.setBancoDados("test_db")

' Now try to build the connection string
conecta.setStringConexao()

Dim connString
connString = conecta.getStringConexao()

Response.Write "Test2-ConexaoSim: sgbd=" & conecta.setSgbd & " result=" & connString & vbCrLf

' Test 3: Mixed Dim/Public/Private
Class MixedClass
    Dim privateVar1
    Private privateVar2
    Public publicVar1
    Dim privateVar3
    
    Public Sub test()
        privateVar1 = "private1"
        privateVar2 = "private2"
        publicVar1 = "public1"
        privateVar3 = "private3"
    End Sub
    
    Public Function getPrivateVar1()
        getPrivateVar1 = privateVar1
    End Function
    
    Public Function getPrivateVar2()
        getPrivateVar2 = privateVar2
    End Function
    
    Public Function getPrivateVar3()
        getPrivateVar3 = privateVar3
    End Function
End Class

Dim obj3
Set obj3 = New MixedClass
obj3.test()
Response.Write "Test3-Mixed: private1=" & obj3.getPrivateVar1() & " private2=" & obj3.getPrivateVar2() & " private3=" & obj3.getPrivateVar3() & " public1=" & obj3.publicVar1 & vbCrLf

Response.Write "=== ALL TESTS COMPLETE ===" & vbCrLf
%>
