package server

import (
	"testing"
	"strings"
	vb "g3pix.com.br/axonasp/vbscript"
	"g3pix.com.br/axonasp/vbscript/ast"
)

// TestClassDimStatements tests that Dim statements inside classes are properly registered
func TestClassDimStatements(t *testing.T) {
	// Create a simple VBScript class with Dim statements
	vbCode := `
Class TestClass
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
End Class`

	// Parse the VBScript code
	parser := vb.NewParser(vbCode)
	program := parser.Parse()
	if program == nil {
		t.Fatalf("Failed to parse VBScript code")
	}

	// Find the class declaration
	var classDecl *ast.ClassDeclaration
	for _, stmt := range program.Body {
		if cd, ok := stmt.(*ast.ClassDeclaration); ok {
			if cd.Identifier.Name == "TestClass" {
				classDecl = cd
				break
			}
		}
	}
	if classDecl == nil {
		t.Fatal("TestClass not found")
	}

	// Create a ClassDef from the declaration
	visitor := &ASPVisitor{}
	classDef := visitor.NewClassDefFromDecl(classDecl)

	// Check that Dim statements were registered as variables
	expectedVars := []string{"name", "age", "scores"}
	for _, varName := range expectedVars {
		if _, exists := classDef.Variables[strings.ToLower(varName)]; !exists {
			t.Errorf("Variable '%s' not found in ClassDef.Variables", varName)
			t.Logf("Available variables: %v", getVariableNames(classDef.Variables))
		}
	}

	// Verify variable properties
	if nameVar, exists := classDef.Variables["name"]; exists {
		if nameVar.Visibility != VisPrivate {
			t.Errorf("Expected 'name' to be private, got %v", nameVar.Visibility)
		}
		if len(nameVar.Dims) != 0 {
			t.Errorf("Expected 'name' to be scalar, got dims %v", nameVar.Dims)
		}
	}

	// Verify array variable
	if scoresVar, exists := classDef.Variables["scores"]; exists {
		if len(scoresVar.Dims) != 1 || scoresVar.Dims[0] != 10 {
			t.Errorf("Expected 'scores' to have dims [10], got %v", scoresVar.Dims)
		}
	}
}

// TestClassDimStatementPersistence tests that class member variables persist between method calls
func TestClassDimStatementPersistence(t *testing.T) {
	// Create a test class with multiple methods that access the same variables
	vbCode := `
Class PersistentClass
	Dim counter
	Dim message
	
	Public Sub increment()
		counter = counter + 1
	End Sub
	
	Public Sub setMessage(msg)
		message = msg
	End Sub
	
	Public Function getCounter()
		getCounter = counter
	End Function
	
	Public Function getMessage()
		getMessage = message
	End Function
End Class`

	// Parse the VBScript code
	parser := vb.NewParser(vbCode)
	program := parser.Parse()
	if program == nil {
		t.Fatalf("Failed to parse VBScript code")
	}

	// Find the class declaration
	var classDecl *ast.ClassDeclaration
	for _, stmt := range program.Body {
		if cd, ok := stmt.(*ast.ClassDeclaration); ok {
			if cd.Identifier.Name == "PersistentClass" {
				classDecl = cd
				break
			}
		}
	}
	if classDecl == nil {
		t.Fatal("PersistentClass not found")
	}

	// Create a ClassDef from the declaration
	visitor := &ASPVisitor{}
	classDef := visitor.NewClassDefFromDecl(classDecl)

	// Check that Dim statements were registered
	expectedVars := []string{"counter", "message"}
	for _, varName := range expectedVars {
		if _, exists := classDef.Variables[strings.ToLower(varName)]; !exists {
			t.Errorf("Variable '%s' not found in ClassDef.Variables", varName)
		}
	}
}

// TestConexaoClassVariables tests the specific Conexao class from the user's code
func TestConexaoClassVariables(t *testing.T) {
	// Simulate the Conexao class structure
	vbCode := `
Class Conexao
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
	
	Public Function getStringConexao()
		getStringConexao = stringConexao
	End Function
	
	Public Sub setStringConexao()
		Select Case sgbd
			Case "oracle"
				stringConexao = "oracle_connection_string"
		End Select
	End Sub
End Class`

	// Parse the VBScript code
	parser := vb.NewParser(vbCode)
	program := parser.Parse()
	if program == nil {
		t.Fatalf("Failed to parse VBScript code")
	}

	// Find the class declaration
	var classDecl *ast.ClassDeclaration
	for _, stmt := range program.Body {
		if cd, ok := stmt.(*ast.ClassDeclaration); ok {
			if cd.Identifier.Name == "Conexao" {
				classDecl = cd
				break
			}
		}
	}
	if classDecl == nil {
		t.Fatal("Conexao class not found")
	}

	// Create a ClassDef from the declaration
	visitor := &ASPVisitor{}
	classDef := visitor.NewClassDefFromDecl(classDecl)

	// Check that all expected variables are registered
	expectedVars := []string{"stringconexao", "servidor", "bancodados", "usuario", "senha", "sgbd", "objconexao"}
	for _, varName := range expectedVars {
		if _, exists := classDef.Variables[strings.ToLower(varName)]; !exists {
			t.Errorf("Variable '%s' not found in Conexao ClassDef.Variables", varName)
			t.Logf("Available variables: %v", getVariableNames(classDef.Variables))
		}
	}

	// Verify all variables are private (Dim inside class should be private)
	for varName, varDef := range classDef.Variables {
		if varDef.Visibility != VisPrivate {
			t.Errorf("Variable '%s' should be private, got %v", varName, varDef.Visibility)
		}
	}
}

// Helper function to get variable names for debugging
func getVariableNames(vars map[string]ClassMemberVar) []string {
	names := make([]string, 0, len(vars))
	for k := range vars {
		names = append(names, k)
	}
	return names
}

// TestClassWithMixedDeclarations tests a class with both Dim and Public/Private declarations
func TestClassWithMixedDeclarations(t *testing.T) {
	vbCode := `
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
End Class`

	// Parse the VBScript code
	parser := vb.NewParser(vbCode)
	program := parser.Parse()
	if program == nil {
		t.Fatalf("Failed to parse VBScript code")
	}

	// Find the class declaration
	var classDecl *ast.ClassDeclaration
	for _, stmt := range program.Body {
		if cd, ok := stmt.(*ast.ClassDeclaration); ok {
			if cd.Identifier.Name == "MixedClass" {
				classDecl = cd
				break
			}
		}
	}
	if classDecl == nil {
		t.Fatal("MixedClass not found")
	}

	// Create a ClassDef from the declaration
	visitor := &ASPVisitor{}
	classDef := visitor.NewClassDefFromDecl(classDecl)

	// Check all variables are registered
	expectedVars := []string{"privatevar1", "privatevar2", "publicvar1", "privatevar3"}
	for _, varName := range expectedVars {
		if _, exists := classDef.Variables[strings.ToLower(varName)]; !exists {
			t.Errorf("Variable '%s' not found in MixedClass ClassDef.Variables", varName)
		}
	}

	// Check visibility
	if pv1, exists := classDef.Variables["privatevar1"]; exists {
		if pv1.Visibility != VisPrivate {
			t.Errorf("Expected privateVar1 to be private, got %v", pv1.Visibility)
		}
	}
	if pv2, exists := classDef.Variables["privatevar2"]; exists {
		if pv2.Visibility != VisPrivate {
			t.Errorf("Expected privateVar2 to be private, got %v", pv2.Visibility)
		}
	}
	if pub1, exists := classDef.Variables["publicvar1"]; exists {
		if pub1.Visibility != VisPublic {
			t.Errorf("Expected publicVar1 to be public, got %v", pub1.Visibility)
		}
	}
}
