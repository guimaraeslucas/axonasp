package vbscript

import (
	"fmt"
	"g3pix.com.br/axonasp/vbscript/ast"
	"strings"
	"testing"
)

func TestASPParser(t *testing.T) {
	code := `
<html>
<body>
<!--#include file="header.asp"-->
<script runat="server">
dim scriptVar
scriptVar = "Script"
</script>
<%
dim name
name = "World"
%>
<h1>Hello <%= name %>!</h1>
<% if name = "World" then %>
  <p>The name is World</p>
<% end if %>
<p>Script var: <%= scriptVar %></p>
<!--#include virtual="/footer.asp"-->
</body>
</html>
`
	parser := NewASPParser(code)
	program := parser.Parse()

	if program == nil {
		t.Fatal("Program is nil")
	}

	for _, stmt := range program.Body {
		printStmt(stmt, 0)
	}
}

func printStmt(stmt ast.Statement, indent int) {
	if stmt == nil {
		return
	}
	var prefix strings.Builder
	for range indent {
		prefix.WriteString("  ")
	}

	switch s := stmt.(type) {
	case *ast.HTMLStatement:
		fmt.Printf("%sHTML: %q\n", prefix.String(), s.Content)
	case *ast.ASPExpressionStatement:
		fmt.Printf("%sASP Expression: %T\n", prefix.String(), s.Expression)
	case *ast.VariablesDeclaration:
		fmt.Printf("%sDim statement\n", prefix.String())
	case *ast.AssignmentStatement:
		fmt.Printf("%sAssignment\n", prefix.String())
	case *ast.IfStatement:
		fmt.Printf("%sIf statement\n", prefix.String())
		if list, ok := s.Consequent.(*ast.StatementList); ok {
			for i := 0; i < list.Count(); i++ {
				printStmt(list.Get(i), indent+1)
			}
		} else if s.Consequent != nil {
			printStmt(s.Consequent, indent+1)
		}
		if s.Alternate != nil {
			fmt.Printf("%sElse\n", prefix.String())
			printStmt(s.Alternate, indent+1)
		}
	case *ast.ElseIfStatement:
		fmt.Printf("%sElseIf\n", prefix.String())
		if list, ok := s.Consequent.(*ast.StatementList); ok {
			for i := 0; i < list.Count(); i++ {
				printStmt(list.Get(i), indent+1)
			}
		} else if s.Consequent != nil {
			printStmt(s.Consequent, indent+1)
		}
		if s.Alternate != nil {
			printStmt(s.Alternate, indent)
		}
	case *ast.ASPDirectiveStatement:
		fmt.Printf("%sASP Directive: %v\n", prefix.String(), s.Attributes)
	case *ast.IncludeStatement:
		fmt.Printf("%sInclude: %s (virtual=%v)\n", prefix.String(), s.Path, s.Virtual)
	default:
		fmt.Printf("%sStatement: %T\n", prefix.String(), s)
	}
}
