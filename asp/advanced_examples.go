package asp

import (
	"fmt"
	"strings"
)

// AdvancedUsageExamples contém exemplos de uso avançado do módulo ASP

// Example1_BasicParsing mostra parsing básico
func Example1_BasicParsing() {
	aspCode := `
<% 
	Dim greeting
	greeting = "Hello, ASP!"
	Response.Write(greeting)
%>
<html>
	<body>
		<p>Welcome</p>
	</body>
</html>
`

	parser := NewASPParser(aspCode)
	result, err := parser.Parse()
	
	if err != nil {
		fmt.Println("Error:", err)
		return
	}
	
	fmt.Println("Example 1: Basic Parsing")
	fmt.Printf("Blocks found: %d\n", len(result.Blocks))
	fmt.Printf("VB Programs: %d\n", len(result.VBPrograms))
	fmt.Println()
}

// Example2_BlockExtraction mostra extração de blocos
func Example2_BlockExtraction() {
	aspCode := `
<% Dim user %>
<h1>User Portal</h1>
<% Response.Write(user) %>
<footer>Copyright 2024</footer>
`

	parser := NewASPParser(aspCode)
	result, _ := parser.Parse()
	
	fmt.Println("Example 2: Block Extraction")
	
	htmlBlocks := 0
	aspBlocks := 0
	
	for i, block := range result.Blocks {
		if block.Type == "html" {
			htmlBlocks++
			fmt.Printf("Block %d (HTML): %s\n", i, strings.TrimSpace(block.Content)[:20]+"...")
		} else if block.Type == "asp" {
			aspBlocks++
			fmt.Printf("Block %d (ASP): %s\n", i, strings.TrimSpace(block.Content))
		}
	}
	
	fmt.Printf("Total HTML blocks: %d\n", htmlBlocks)
	fmt.Printf("Total ASP blocks: %d\n", aspBlocks)
	fmt.Println()
}

// Example3_CodeExtraction mostra extração separada de código
func Example3_CodeExtraction() {
	aspCode := `
<% 
	Dim conn
	Set conn = CreateObject("ADODB.Connection")
%>
<html>
	<body>Database Connected</body>
</html>
<% conn.Close %>
`

	html := ExtractHTMLOnly(aspCode)
	vb := ExtractVBScriptOnly(aspCode)
	
	fmt.Println("Example 3: Code Extraction")
	fmt.Println("HTML Content:")
	fmt.Println(html)
	fmt.Println("\nVBScript Content:")
	fmt.Println(vb)
	fmt.Println()
}

// Example4_Validation mostra validação de código
func Example4_Validation() {
	testCases := []struct {
		name string
		code string
	}{
		{"Valid code", "<% Dim x %>"},
		{"HTML only", "<html></html>"},
		{"Complex ASP", "<% Dim x %>\n<html>\n<% Response.Write(x) %></html>"},
	}
	
	fmt.Println("Example 4: Code Validation")
	
	for _, tc := range testCases {
		valid, errors := ValidateASP(tc.code)
		fmt.Printf("%s: %v", tc.name, valid)
		if len(errors) > 0 {
			fmt.Printf(" - Error: %v", errors[0])
		}
		fmt.Println()
	}
	fmt.Println()
}

// Example5_ASPObjects mostra uso de objetos ASP
func Example5_ASPObjects() {
	fmt.Println("Example 5: ASP Objects")
	
	// Cria contexto ASP
	ctx := NewASPContext()
	
	// Usa objeto Server
	encoded, _ := ctx.Server.CallMethod("URLEncode", "hello world & special?")
	fmt.Printf("URLEncoded: %s\n", encoded)
	
	htmlEncoded, _ := ctx.Server.CallMethod("HTMLEncode", "<script>alert('xss')</script>")
	fmt.Printf("HTMLEncoded: %s\n", htmlEncoded)
	
	// Usa objeto Response
	ctx.Response.CallMethod("Write", "First line\n")
	ctx.Response.CallMethod("Write", "Second line\n")
	fmt.Printf("Response buffer: %s\n", ctx.Response.GetBuffer())
	
	// Usa objeto Session
	ctx.Session.SetProperty("userid", 12345)
	ctx.Session.SetProperty("username", "john")
	fmt.Printf("Session userid: %v\n", ctx.Session.GetProperty("userid"))
	fmt.Println()
}

// Example6_ErrorHandling mostra tratamento de erros
func Example6_ErrorHandling() {
	fmt.Println("Example 6: Error Handling")
	
	aspCode := `
<%
	Dim file
	Set file = CreateObject("Scripting.FileSystemObject")
	
	If Err.Number <> 0 Then
		Response.Write("Error: " & Err.Description)
	End If
%>
`
	
	parser := NewASPParser(aspCode)
	result, err := parser.Parse()
	
	if err != nil {
		fmt.Printf("Critical error: %v\n", err)
	}
	
	if len(result.Errors) > 0 {
		fmt.Printf("Found %d parse errors:\n", len(result.Errors))
		for _, e := range result.Errors {
			fmt.Printf("  - %v\n", e)
		}
	} else {
		fmt.Println("No parse errors found")
	}
	fmt.Println()
}

// Example7_Formatting mostra formatação de código
func Example7_Formatting() {
	aspCode := `
<%Dim x%>
<html>
<%Response.Write("test")%>
</html>
`

	formatter := NewASPFormatter(2)
	formatted := formatter.Format(aspCode)
	
	fmt.Println("Example 7: Code Formatting")
	fmt.Println("Original:")
	fmt.Println(aspCode)
	fmt.Println("\nFormatted:")
	fmt.Println(formatted)
	fmt.Println()
}

// Example8_Analysis mostra análise de código
func Example8_Analysis() {
	aspCode := `
<%
	Dim conn, rs
	Set conn = CreateObject("ADODB.Connection")
	conn.Open "Provider=SQLOLEDB;Data Source=myserver;"
	
	Set rs = conn.Execute("SELECT * FROM users WHERE id = " & Request.QueryString("id"))
	
	If Not rs.EOF Then
		Response.Write("Found user: " & rs("name"))
	End If
	
	rs.Close
	conn.Close
	Set rs = Nothing
	Set conn = Nothing
%>
<html>
	<body>
		<h1>User Details</h1>
		<p>User info loaded</p>
	</body>
</html>
`

	analyzer := NewASPCodeAnalyzer()
	analysis := analyzer.Analyze(aspCode)
	
	fmt.Println("Example 8: Code Analysis")
	fmt.Printf("Total blocks: %d\n", analysis["total_blocks"])
	fmt.Printf("ASP blocks: %d\n", analysis["asp_blocks"])
	fmt.Printf("HTML blocks: %d\n", analysis["html_blocks"])
	fmt.Printf("Complexity: %s\n", analysis["complexity"])
	
	if patterns, ok := analysis["patterns_detected"].([]string); ok {
		fmt.Println("Patterns detected:")
		for _, p := range patterns {
			fmt.Printf("  - %s\n", p)
		}
	}
	fmt.Println()
}

// Example9_Integration mostra integração com VBScript-Go
func Example9_Integration() {
	aspCode := `
<% 
	Dim message
	message = "Integrated with VBScript-Go"
%>
<html>
	<body>
		<% Response.Write(message) %>
	</body>
</html>
`

	integration := NewASPVBIntegration()
	result, _ := integration.ParseASPFile("example.asp", aspCode)
	vbCode := integration.GetVBScriptCodeForVBScriptGo()
	
	fmt.Println("Example 9: VBScript-Go Integration")
	fmt.Printf("Parse result blocks: %d\n", len(result.Blocks))
	fmt.Printf("Extracted VB code length: %d\n", len(vbCode))
	
	info := integration.GetCompatibilityInfo()
	fmt.Printf("Compatible: %v\n", info["vbscript_compatible"])
	
	if objs, ok := info["asp_objects_supported"].([]string); ok {
		fmt.Println("ASP Objects supported:")
		for _, obj := range objs {
			fmt.Printf("  - %s\n", obj)
		}
	}
	fmt.Println()
}

// Example10_CompleteWorkflow mostra workflow completo
func Example10_CompleteWorkflow() {
	aspFile := `
<%
	' Login example
	If Request.Form("action") = "login" Then
		Dim username, password
		username = Request.Form("username")
		password = Request.Form("password")
		
		If Len(username) > 0 Then
			Session("authenticated") = True
			Session("username") = username
			Response.Redirect("dashboard.asp")
		End If
	End If
%>

<!DOCTYPE html>
<html>
<head>
	<title>Login</title>
</head>
<body>
	<h1>User Login</h1>
	<form method="post">
		<input type="hidden" name="action" value="login">
		<label>Username:</label>
		<input type="text" name="username" required>
		<label>Password:</label>
		<input type="password" name="password" required>
		<input type="submit" value="Login">
	</form>
</body>
</html>

<%
	' Footer
	Response.Write("<!-- Page generated at " & Now() & " -->")
%>
`

	fmt.Println("Example 10: Complete Workflow")
	
	// 1. Parse
	parser := NewASPParser(aspFile)
	result, _ := parser.Parse()
	fmt.Printf("1. Parsed %d blocks\n", len(result.Blocks))
	
	// 2. Analyze
	analyzer := NewASPCodeAnalyzer()
	analysis := analyzer.Analyze(aspFile)
	fmt.Printf("2. Complexity: %s\n", analysis["complexity"])
	
	// 3. Extract components
	html := ExtractHTMLOnly(aspFile)
	vb := ExtractVBScriptOnly(aspFile)
	fmt.Printf("3. Extracted HTML (%d bytes) and VB (%d bytes)\n", len(html), len(vb))
	
	// 4. Validate
	valid, _ := ValidateASP(aspFile)
	fmt.Printf("4. Validation: %v\n", valid)
	
	// 5. Format
	formatter := NewASPFormatter(2)
	formatted := formatter.Format(aspFile)
	fmt.Printf("5. Formatted %d bytes\n", len(formatted))
	
	fmt.Println("\nWorkflow completed successfully!")
	fmt.Println()
}

// RunAllExamples executa todos os exemplos
func RunAllExamples() {
	fmt.Println("====================================")
	fmt.Println("ASP Module - Advanced Usage Examples")
	fmt.Println("====================================\n")
	
	Example1_BasicParsing()
	Example2_BlockExtraction()
	Example3_CodeExtraction()
	Example4_Validation()
	Example5_ASPObjects()
	Example6_ErrorHandling()
	Example7_Formatting()
	Example8_Analysis()
	Example9_Integration()
	Example10_CompleteWorkflow()
	
	fmt.Println("====================================")
	fmt.Println("All examples completed!")
	fmt.Println("====================================")
}
