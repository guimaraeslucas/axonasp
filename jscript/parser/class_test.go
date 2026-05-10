package parser

import (
	"g3pix.com.br/axonasp/jscript/ast"
	"testing"
)

func TestClassParsing(t *testing.T) {
	src := `class A extends B { constructor() { super(); } method() {} }`
	program, err := ParseFile(nil, "", src, 0)
	if err != nil {
		t.Fatalf("Parse error: %v", err)
	}

	if len(program.Body) != 1 {
		t.Fatalf("Expected 1 statement, got %d", len(program.Body))
	}

	stmt := program.Body[0]
	classDecl, ok := stmt.(*ast.ClassDeclaration)
	if !ok {
		t.Fatalf("Expected ClassDeclaration, got %T", stmt)
	}

	class := classDecl.Class
	if class.Name.Name != "A" {
		t.Errorf("Expected class name A, got %s", class.Name.Name)
	}

	superClass, ok := class.SuperClass.(*ast.Identifier)
	if !ok || superClass.Name != "B" {
		t.Errorf("Expected super class B, got %v", class.SuperClass)
	}

	if len(class.Body) != 2 {
		t.Fatalf("Expected 2 elements in class body, got %d", len(class.Body))
	}

	// Check constructor
	ctor, ok := class.Body[0].(*ast.MethodDefinition)
	if !ok || ctor.Kind != ast.PropertyKindConstructor {
		t.Errorf("Expected constructor, got %v", class.Body[0])
	}

	// Check super() call in constructor
	// ctor.Body is *FunctionLiteral
	// ctor.Body.Body is *BlockStatement
	ctorBlock := ctor.Body.Body
	if ctorBlock == nil {
		t.Fatalf("Expected BlockStatement for constructor body, got nil")
	}
	if len(ctorBlock.List) != 1 {
		t.Fatalf("Expected 1 statement in constructor, got %d", len(ctorBlock.List))
	}
	exprStmt, ok := ctorBlock.List[0].(*ast.ExpressionStatement)
	if !ok {
		t.Fatalf("Expected ExpressionStatement, got %T", ctorBlock.List[0])
	}
	callExpr, ok := exprStmt.Expression.(*ast.CallExpression)
	if !ok {
		t.Fatalf("Expected CallExpression, got %T", exprStmt.Expression)
	}
	_, ok = callExpr.Callee.(*ast.SuperExpression)
	if !ok {
		t.Errorf("Expected super() call, got callee %T", callExpr.Callee)
	}

	// Check method
	method, ok := class.Body[1].(*ast.MethodDefinition)
	if !ok || method.Kind != ast.PropertyKindMethod {
		t.Errorf("Expected method, got %v", class.Body[1])
	}
}

func TestClassExpressionParsing(t *testing.T) {
	src := `var x = class C {}`
	program, err := ParseFile(nil, "", src, 0)
	if err != nil {
		t.Fatalf("Parse error: %v", err)
	}

	// program -> VariableStatement -> []*Binding -> Binding -> ClassExpression
	stmt := program.Body[0].(*ast.VariableStatement)
	binding := stmt.List[0]
	classExpr, ok := binding.Initializer.(*ast.ClassExpression)
	if !ok {
		t.Fatalf("Expected ClassExpression, got %T", binding.Initializer)
	}

	if classExpr.Class.Name.Name != "C" {
		t.Errorf("Expected class name C, got %s", classExpr.Class.Name.Name)
	}
}

func TestStaticMethodParsing(t *testing.T) {
	src := `class A { static stMethod() {} }`
	program, err := ParseFile(nil, "", src, 0)
	if err != nil {
		t.Fatalf("Parse error: %v", err)
	}

	classDecl := program.Body[0].(*ast.ClassDeclaration)
	class := classDecl.Class
	method := class.Body[0].(*ast.MethodDefinition)
	if !method.Static {
		t.Error("Expected method to be static")
	}
}

func TestGetSetParsing(t *testing.T) {
	src := `class A { get x() {} set x(v) {} }`
	program, err := ParseFile(nil, "", src, 0)
	if err != nil {
		t.Fatalf("Parse error: %v", err)
	}

	classDecl := program.Body[0].(*ast.ClassDeclaration)
	class := classDecl.Class
	getProp := class.Body[0].(*ast.MethodDefinition)
	if getProp.Kind != ast.PropertyKindGet {
		t.Errorf("Expected get, got %s", getProp.Kind)
	}
	setProp := class.Body[1].(*ast.MethodDefinition)
	if setProp.Kind != ast.PropertyKindSet {
		t.Errorf("Expected set, got %s", setProp.Kind)
	}
}
