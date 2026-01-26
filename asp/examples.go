/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * Attribution Notice:
 * If this software is used in other projects, the name "AxonASP Server"
 * must be cited in the documentation or "About" section.
 *
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 */
package asp

import (
	"strings"
)

// ExampleASPCode contém exemplos de código ASP para testes
var ExampleASPCode = map[string]string{
	"simple": `
<html>
<body>
	<% Response.Write("Hello World") %>
</body>
</html>
`,

	"multiple_blocks": `
<html>
<body>
	<% 
		Dim name
		name = "John"
	%>
	
	<h1>Welcome <%= name %></h1>
	
	<% 
		Response.Write("Current time: " & Now())
	%>
</body>
</html>
`,

	"with_database": `
<%
	Dim conn, rs
	Set conn = CreateObject("ADODB.Connection")
	conn.Open "Provider=SQLOLEDB;Data Source=myserver;Initial Catalog=mydb;"
	
	Set rs = conn.Execute("SELECT * FROM users")
%>

<table>
	<tr>
		<th>ID</th>
		<th>Name</th>
	</tr>
	<% 
		Do While Not rs.EOF
			Response.Write("<tr><td>" & rs("id") & "</td><td>" & rs("name") & "</td></tr>")
			rs.MoveNext
		Loop
	%>
</table>

<%
	rs.Close
	conn.Close
	Set rs = Nothing
	Set conn = Nothing
%>
`,

	"with_form": `
<%
	If Request.Form("submit") <> "" Then
		Dim username, password
		username = Request.Form("username")
		password = Request.Form("password")
		Response.Write("Login attempt: " & username)
	End If
%>

<form method="post">
	<input type="text" name="username" placeholder="Username">
	<input type="password" name="password" placeholder="Password">
	<input type="submit" name="submit" value="Login">
</form>
`,

	"complex": `
<%
	Option Explicit
	
	Dim server_name, db_path
	server_name = Request.ServerVariables("SERVER_NAME")
	db_path = Server.MapPath("/data/database.mdb")
	
	Function ValidateInput(input)
		ValidateInput = Len(input) > 0
	End Function
	
	Sub WriteLog(message)
		' Log implementation here
	End Sub
	
	If Request.Form("action") = "login" Then
		If ValidateInput(Request.Form("user")) Then
			Session("authenticated") = True
			Response.Redirect("/dashboard.asp")
		End If
	End If
%>

<!DOCTYPE html>
<html>
<head>
	<title>Login</title>
</head>
<body>
	<%
		If Not Session("authenticated") Then
	%>
		<form method="post">
			<input type="hidden" name="action" value="login">
			<label>Username:</label>
			<input type="text" name="user">
			<input type="submit" value="Login">
		</form>
	<%
		Else
	%>
		<p>Welcome back, <%= Session("username") %></p>
	<%
		End If
	%>
</body>
</html>
`,

	"json_example": `
<%
	Response.ContentType = "application/json"
	
	Dim data
	data = "{"
	data = data & """name"": ""John"","
	data = data & """age"": 30,"
	data = data & """active"": true"
	data = data & "}"
	
	Response.Write(data)
%>
`,

	"error_handling": `
<%
	On Error Resume Next
	
	Dim file
	Set file = CreateObject("Scripting.FileSystemObject")
	
	If Err.Number <> 0 Then
		Response.Write("Error: " & Err.Description)
		Err.Clear
	End If
%>

<html>
<body>
	<p>Processing request...</p>
</body>
</html>
`,

	"session_example": `
<%
	If Request.QueryString("logout") <> "" Then
		Session.Abandon
		Response.Redirect("/index.asp")
	End If
	
	If Session("logged_in") = "" Then
		Session("login_attempts") = 0
	End If
%>

<html>
<body>
	<% If Session("logged_in") = True Then %>
		<p>You are logged in as: <%= Session("username") %></p>
		<a href="?logout=1">Logout</a>
	<% Else %>
		<p>Please log in first</p>
	<% End If %>
</body>
</html>
`,

	"include_example": `
<!-- #include file="header.inc" -->

<%
	Dim title
	title = "My Page"
	Response.Write("<h1>" & title & "</h1>")
%>

<!-- #include file="footer.inc" -->
`,
}

// TestCodeBlock testa um bloco de código específico
func TestCodeBlock(codeType string) (bool, string) {
	code, exists := ExampleASPCode[codeType]
	if !exists {
		return false, "Code type not found: " + codeType
	}

	valid, errors := ValidateASP(code)

	if !valid && len(errors) > 0 {
		return false, errors[0].Error()
	}

	return valid, "OK"
}

// GetAllExamples retorna todos os exemplos disponíveis
func GetAllExamples() []string {
	var examples []string
	for key := range ExampleASPCode {
		examples = append(examples, key)
	}
	return examples
}

// GetExampleCode retorna o código de um exemplo específico
func GetExampleCode(codeType string) string {
	if code, exists := ExampleASPCode[codeType]; exists {
		return code
	}
	return ""
}

// AnalyzeExample analisa um exemplo e retorna detalhes
func AnalyzeExample(codeType string) string {
	code, exists := ExampleASPCode[codeType]
	if !exists {
		return "Example not found"
	}

	parser := NewASPParser(code)
	result, err := parser.Parse()

	output := strings.Builder{}
	output.WriteString("=== Analysis of: " + codeType + " ===\n")

	if err != nil {
		output.WriteString("Parse Error: " + err.Error() + "\n")
		return output.String()
	}

	output.WriteString("\nTotal Blocks: ")
	output.WriteString(string(rune(len(result.Blocks))))
	output.WriteString("\n")

	htmlCount := 0
	aspCount := 0

	for _, block := range result.Blocks {
		if block.Type == "html" {
			htmlCount++
		} else if block.Type == "asp" {
			aspCount++
		}
	}

	output.WriteString("HTML Blocks: ")
	output.WriteString(string(rune(htmlCount)))
	output.WriteString("\n")

	output.WriteString("ASP Blocks: ")
	output.WriteString(string(rune(aspCount)))
	output.WriteString("\n")

	if len(result.Errors) > 0 {
		output.WriteString("\nParse Errors:\n")
		for _, parseErr := range result.Errors {
			output.WriteString("  - " + parseErr.Error() + "\n")
		}
	}

	return output.String()
}
