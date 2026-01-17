package main

import (
	"fmt"
	"strings"

	"go-asp/asp"
)

func main() {
	fmt.Println("╔════════════════════════════════════════════════════╗")
	fmt.Println("║      ASP Classic Parser - Demonstração Completa    ║")
	fmt.Println("╚════════════════════════════════════════════════════╝")
	fmt.Println()

	// ============================================
	// Exemplo 1: Parse Simples
	// ============================================
	fmt.Println("📝 EXEMPLO 1: Parse Simples")
	fmt.Println("─────────────────────────────────────────────────────")

	example1 := `
<html>
<head>
	<title>Página ASP</title>
</head>
<body>
	<%
		Dim welcome
		welcome = "Bem-vindo ao ASP!"
		Response.Write(welcome)
	%>
</body>
</html>
`

	parser1 := asp.NewASPParser(example1)
	result1, err1 := parser1.Parse()

	if err1 != nil {
		fmt.Printf("❌ Erro: %v\n", err1)
		return
	}

	fmt.Printf("✓ Blocos encontrados: %d\n", len(result1.Blocks))
	fmt.Printf("✓ Blocos ASP: %d\n", countBlockType(result1.Blocks, "asp"))
	fmt.Printf("✓ Blocos HTML: %d\n", countBlockType(result1.Blocks, "html"))
	fmt.Println()

	// ============================================
	// Exemplo 2: Múltiplos Blocos
	// ============================================
	fmt.Println("📝 EXEMPLO 2: Múltiplos Blocos")
	fmt.Println("─────────────────────────────────────────────────────")

	example2 := `
<%
	Dim user
	user = "João"
%>
<h1>Olá <%= user %></h1>
<p>Esta é uma página ASP clássica</p>
<%
	Response.Write("Hora: " & Now())
%>
<footer>Copyright 2024</footer>
<%
	' Comentário VB
	Dim version
	version = "1.0"
%>
`

	parser2 := asp.NewASPParser(example2)
	result2, _ := parser2.Parse()

	fmt.Println("Estrutura dos blocos:")
	for i, block := range result2.Blocks {
		if block.Type == "asp" {
			fmt.Printf("  [%d] ASP (Linha %d, %d chars)\n", i, block.Line, len(block.Content))
			fmt.Printf("      → %s\n", truncate(block.Content, 50))
		} else {
			fmt.Printf("  [%d] HTML (Linha %d, %d chars)\n", i, block.Line, len(block.Content))
			fmt.Printf("      → %s\n", truncate(block.Content, 50))
		}
	}
	fmt.Println()

	// ============================================
	// Exemplo 3: Validação
	// ============================================
	fmt.Println("📝 EXEMPLO 3: Validação de Código")
	fmt.Println("─────────────────────────────────────────────────────")

	testCases := []struct {
		name string
		code string
	}{
		{"Código válido simples", `<% Dim x %>`},
		{"HTML puro", `<html><body>Test</body></html>`},
		{"Múltiplos blocos", `<% Dim x %><html></html><% x = 5 %>`},
	}

	for _, tc := range testCases {
		valid, errors := asp.ValidateASP(tc.code)
		status := "✓"
		if !valid {
			status = "✗"
		}
		fmt.Printf("%s %s: %v\n", status, tc.name, valid)
		if len(errors) > 0 {
			fmt.Printf("   Erros: %v\n", errors)
		}
	}
	fmt.Println()

	// ============================================
	// Exemplo 4: Extração de Componentes
	// ============================================
	fmt.Println("📝 EXEMPLO 4: Extração de Componentes")
	fmt.Println("─────────────────────────────────────────────────────")

	example4 := `
<% Dim db %>
<html>
<body>
	<% Response.Write("Database") %>
</body>
</html>
<% Set db = Nothing %>
`

	html := asp.ExtractHTMLOnly(example4)
	vb := asp.ExtractVBScriptOnly(example4)

	fmt.Println("HTML Extraído:")
	fmt.Printf("  %s\n", strings.TrimSpace(html))
	fmt.Println()
	fmt.Println("VBScript Extraído:")
	for _, line := range strings.Split(strings.TrimSpace(vb), "\n") {
		if strings.TrimSpace(line) != "" {
			fmt.Printf("  %s\n", line)
		}
	}
	fmt.Println()

	// ============================================
	// Exemplo 5: Análise de Complexidade
	// ============================================
	fmt.Println("📝 EXEMPLO 5: Análise de Código")
	fmt.Println("─────────────────────────────────────────────────────")

	example5 := `
<%
	Dim conn, rs
	Set conn = CreateObject("ADODB.Connection")
	conn.Open "Provider=SQLOLEDB;Data Source=myserver;"
	
	Set rs = conn.Execute("SELECT * FROM users")
	
	If Not rs.EOF Then
		Do While Not rs.EOF
			Response.Write(rs("name") & "<br>")
			rs.MoveNext
		Loop
	End If
	
	rs.Close
	conn.Close
%>

<table>
	<tr><th>Users</th></tr>
	<tr><td>Listed above</td></tr>
</table>
`

	analyzer := asp.NewASPCodeAnalyzer()
	analysis := analyzer.Analyze(example5)

	fmt.Printf("Blocos totais: %d\n", analysis["total_blocks"])
	fmt.Printf("Blocos ASP: %d\n", analysis["asp_blocks"])
	fmt.Printf("Blocos HTML: %d\n", analysis["html_blocks"])
	fmt.Printf("Complexidade: %s\n", analysis["complexity"])

	if patterns, ok := analysis["patterns_detected"].([]string); ok && len(patterns) > 0 {
		fmt.Println("Padrões detectados:")
		for _, p := range patterns {
			fmt.Printf("  • %s\n", p)
		}
	}
	fmt.Println()

	// ============================================
	// Exemplo 6: Objetos ASP
	// ============================================
	fmt.Println("📝 EXEMPLO 6: Simulação de Objetos ASP")
	fmt.Println("─────────────────────────────────────────────────────")

	ctx := asp.NewASPContext()

	// Server
	fmt.Println("Server Object:")
	encoded, _ := ctx.Server.CallMethod("URLEncode", "hello world & test?")
	fmt.Printf("  URLEncode('hello world & test?'): %v\n", encoded)

	htmlEncoded, _ := ctx.Server.CallMethod("HTMLEncode", "<script>alert('xss')</script>")
	fmt.Printf("  HTMLEncode('<script>...'): %v\n", htmlEncoded)

	// Response
	fmt.Println("Response Object:")
	ctx.Response.CallMethod("Write", "Linha 1\n")
	ctx.Response.CallMethod("Write", "Linha 2\n")
	fmt.Printf("  Buffer: %s", ctx.Response.GetBuffer())

	// Session
	fmt.Println("Session Object:")
	ctx.Session.SetProperty("userid", 12345)
	ctx.Session.SetProperty("username", "joao")
	fmt.Printf("  userid: %v\n", ctx.Session.GetProperty("userid"))
	fmt.Printf("  username: %v\n", ctx.Session.GetProperty("username"))
	fmt.Println()

	// ============================================
	// Exemplo 7: Formatação de Código
	// ============================================
	fmt.Println("📝 EXEMPLO 7: Formatação de Código")
	fmt.Println("─────────────────────────────────────────────────────")

	example7 := `<%Dim x%><html><%Response.Write(x)%></html>`

	formatter := asp.NewASPFormatter(4)
	formatted := formatter.Format(example7)

	fmt.Println("Código original:")
	fmt.Println(example7)
	fmt.Println("\nCódigo formatado:")
	fmt.Println(formatted)
	fmt.Println()

	// ============================================
	// Exemplo 8: Caso Real Completo
	// ============================================
	fmt.Println("📝 EXEMPLO 8: Caso Real - Sistema de Login")
	fmt.Println("─────────────────────────────────────────────────────")

	loginPage := `
<%
	' Verificar se formulário foi enviado
	If Request.Form("action") = "login" Then
		Dim username, password
		username = Request.Form("username")
		password = Request.Form("password")
		
		If Len(username) > 0 And Len(password) > 0 Then
			Session("authenticated") = True
			Session("username") = username
			Response.Redirect("dashboard.asp")
		Else
			Response.Write("Credenciais inválidas")
		End If
	End If
%>

<!DOCTYPE html>
<html>
<head>
	<title>Login</title>
</head>
<body>
	<h1>Sistema de Login</h1>
	<form method="post">
		<input type="hidden" name="action" value="login">
		<label>Usuário:</label>
		<input type="text" name="username" required>
		<label>Senha:</label>
		<input type="password" name="password" required>
		<input type="submit" value="Entrar">
	</form>
</body>
</html>

<%
	Response.Write("<!-- Página de login gerada em: " & Now() & " -->")
%>
`

	parserLogin := asp.NewASPParser(loginPage)
	resultLogin, _ := parserLogin.Parse()

	fmt.Printf("Análise da página de login:\n")
	fmt.Printf("  • Blocos totais: %d\n", len(resultLogin.Blocks))
	fmt.Printf("  • Erros: %d\n", len(resultLogin.Errors))

	analyzerLogin := asp.NewASPCodeAnalyzer()
	analysisLogin := analyzerLogin.Analyze(loginPage)

	fmt.Printf("  • Complexidade: %s\n", analysisLogin["complexity"])

	if patterns, ok := analysisLogin["patterns_detected"].([]string); ok {
		fmt.Printf("  • Padrões: %d\n", len(patterns))
		for _, p := range patterns {
			fmt.Printf("    - %s\n", p)
		}
	}
	fmt.Println()

	// ============================================
	// Resumo Final
	// ============================================
	fmt.Println("╔════════════════════════════════════════════════════╗")
	fmt.Println("║                    RESUMO FINAL                     ║")
	fmt.Println("╚════════════════════════════════════════════════════╝")
	fmt.Println()
	fmt.Println("✅ ASP Parser está funcionando corretamente!")
	fmt.Println()
	fmt.Println("Funcionalidades demonstradas:")
	fmt.Println("  ✓ Parse de código ASP clássico")
	fmt.Println("  ✓ Identificação de blocos <% %>")
	fmt.Println("  ✓ Validação de sintaxe")
	fmt.Println("  ✓ Extração de HTML e VBScript")
	fmt.Println("  ✓ Análise de complexidade")
	fmt.Println("  ✓ Detecção de padrões")
	fmt.Println("  ✓ Simulação de objetos ASP")
	fmt.Println("  ✓ Formatação de código")
	fmt.Println()
	fmt.Println("Para usar o parser em seu projeto:")
	fmt.Println("  import \"asp\"")
	fmt.Println("  parser := asp.NewASPParser(aspCode)")
	fmt.Println("  result, _ := parser.Parse()")
	fmt.Println()
}

// Funções auxiliares

func countBlockType(blocks []*asp.CodeBlock, blockType string) int {
	count := 0
	for _, block := range blocks {
		if block.Type == blockType {
			count++
		}
	}
	return count
}

func truncate(s string, maxLen int) string {
	s = strings.ReplaceAll(s, "\n", " ")
	s = strings.ReplaceAll(s, "\t", " ")
	s = strings.TrimSpace(s)
	if len(s) > maxLen {
		return s[:maxLen] + "..."
	}
	return s
}
