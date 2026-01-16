package main

import (
	"fmt"
	"os"
	"strings"

	vbs "github.com/guimaraeslucas/vbscript-go"
	"github.com/guimaraeslucas/vbscript-go/ast"
)

func main() {
	// Exemplo 1: Código VBScript simples inline
	fmt.Println("=== Exemplo 1: Código Inline Simples ===")
	code1 := `Dim x
x = 10`
	parseAndDisplay(code1)

	fmt.Println("\n=== Exemplo 2: Com Option Explicit ===")
	code2 := `Option Explicit
Dim y
y = 20`
	parseAndDisplay(code2)

	fmt.Println("\n=== Exemplo 3: Function e Sub ===")
	code3 := `Function Soma(a, b)
    Soma = a + b
End Function

Sub Exibir()
End Sub`
	parseAndDisplay(code3)

	// Exemplo 4: Ler de arquivo se fornecido
	if len(os.Args) > 1 {
		fmt.Println("\n=== Exemplo 4: Arquivo Fornecido ===")
		filePath := os.Args[1]
		content, err := os.ReadFile(filePath)
		if err != nil {
			fmt.Printf("Erro ao ler arquivo: %v\n", err)
			return
		}
		parseAndDisplay(string(content))
	}
}

func parseAndDisplay(code string) {
	// Criar o parser
	parser := vbs.NewParser(code)

	// Fazer o parsing
	defer func() {
		if r := recover(); r != nil {
			fmt.Printf("❌ Erro de parsing: %v\n", r)
			// Mostrar o código que causou o erro (primeiras 200 chars)
			if len(code) > 200 {
				fmt.Printf("Código (primeiros 200 chars): %s...\n", code[:200])
			} else {
				fmt.Printf("Código: %s\n", code)
			}
		}
	}()

	program := parser.Parse()

	// Exibir informações sobre o programa
	fmt.Printf("✅ Parsing bem-sucedido!\n")
	fmt.Printf("📋 Option Explicit: %v\n", program.OptionExplicit)
	fmt.Printf("📊 Número de statements: %d\n", len(program.Body))
	
	// Percorrer e exibir os statements
	fmt.Println("\n📝 Estrutura do AST:")
	for i, stmt := range program.Body {
		displayStatement(stmt, i+1, 0)
	}
}

func displayStatement(stmt ast.Statement, index int, indent int) {
	indentStr := strings.Repeat("  ", indent)
	
	switch s := stmt.(type) {
	case *ast.VariablesDeclaration:
		fmt.Printf("%s%d. VariablesDeclaration (%d variáveis)\n", indentStr, index, len(s.Variables))
		for _, v := range s.Variables {
			fmt.Printf("%s   - %s\n", indentStr, v.Identifier.Name)
		}
		
	case *ast.AssignmentStatement:
		fmt.Printf("%s%d. AssignmentStatement: ", indentStr, index)
		displayExpression(s.Left)
		fmt.Printf(" = ")
		displayExpression(s.Right)
		fmt.Println()
		
	case *ast.IfStatement:
		fmt.Printf("%s%d. IfStatement\n", indentStr, index)
		fmt.Printf("%s   Test: ", indentStr)
		displayExpression(s.Test)
		fmt.Println()
		fmt.Printf("%s   Consequent: ", indentStr)
		if s.Consequent != nil {
			displayStatement(s.Consequent, 1, indent+2)
		}
		if s.Alternate != nil {
			fmt.Printf("%s   Alternate: ", indentStr)
			displayStatement(s.Alternate, 1, indent+2)
		}
		
	case *ast.ForStatement:
		fmt.Printf("%s%d. ForStatement: %s = ", indentStr, index, s.Identifier.Name)
		displayExpression(s.From)
		fmt.Printf(" To ")
		displayExpression(s.To)
		if s.Step != nil {
			fmt.Printf(" Step ")
			displayExpression(s.Step)
		}
		fmt.Printf(" (%d statements)\n", len(s.Body))
		
	case *ast.WhileStatement:
		fmt.Printf("%s%d. WhileStatement\n", indentStr, index)
		fmt.Printf("%s   Condition: ", indentStr)
		displayExpression(s.Condition)
		fmt.Printf(" (%d statements)\n", len(s.Body))
		
	case *ast.DoStatement:
		fmt.Printf("%s%d. DoStatement (LoopType: %v, TestType: %v)\n", indentStr, index, s.LoopType, s.TestType)
		if s.Condition != nil {
			fmt.Printf("%s   Condition: ", indentStr)
			displayExpression(s.Condition)
			fmt.Println()
		}
		fmt.Printf("%s   (%d statements)\n", indentStr, len(s.Body))
		
	case *ast.SelectStatement:
		fmt.Printf("%s%d. SelectStatement\n", indentStr, index)
		fmt.Printf("%s   Condition: ", indentStr)
		displayExpression(s.Condition)
		fmt.Printf("\n%s   Cases: %d\n", indentStr, len(s.Cases))
		
	case *ast.FunctionDeclaration:
		fmt.Printf("%s%d. FunctionDeclaration: %s (%d parâmetros)\n", 
			indentStr, index, s.Identifier.Name, len(s.Parameters))
		
	case *ast.SubDeclaration:
		fmt.Printf("%s%d. SubDeclaration: %s (%d parâmetros)\n", 
			indentStr, index, s.Identifier.Name, len(s.Parameters))
		
	case *ast.CallStatement:
		fmt.Printf("%s%d. CallStatement: ", indentStr, index)
		displayExpression(s.Callee)
		fmt.Println()
		
	default:
		fmt.Printf("%s%d. %T\n", indentStr, index, stmt)
	}
}

func displayExpression(expr ast.Expression) {
	if expr == nil {
		fmt.Print("<nil>")
		return
	}
	
	switch e := expr.(type) {
	case *ast.Identifier:
		fmt.Print(e.Name)
		
	case *ast.StringLiteral:
		fmt.Printf("\"%s\"", e.Value)
		
	case *ast.IntegerLiteral:
		fmt.Printf("%d", e.Value)
		
	case *ast.FloatLiteral:
		fmt.Printf("%f", e.Value)
		
	case *ast.BooleanLiteral:
		fmt.Printf("%v", e.Value)
		
	case *ast.BinaryExpression:
		fmt.Print("(")
		displayExpression(e.Left)
		fmt.Printf(" %v ", e.Operation)
		displayExpression(e.Right)
		fmt.Print(")")
		
	case *ast.UnaryExpression:
		fmt.Printf("%v(", e.Operation)
		displayExpression(e.Argument)
		fmt.Print(")")
		
	case *ast.MemberExpression:
		displayExpression(e.Object)
		fmt.Print(".")
		displayExpression(e.Property)
		
	case *ast.IndexOrCallExpression:
		displayExpression(e.Object)
		fmt.Print("(")
		for i, arg := range e.Indexes {
			if i > 0 {
				fmt.Print(", ")
			}
			displayExpression(arg)
		}
		fmt.Print(")")
		
	default:
		fmt.Printf("<%T>", expr)
	}
}
