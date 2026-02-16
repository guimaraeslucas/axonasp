package experimental

import (
	"fmt"
	"testing"

	"g3pix.com.br/axonasp/vbscript/ast"
)

func TestSimpleMath(t *testing.T) {
	// 1 + 2 * 3
	// AST: Binary(+, 1, Binary(*, 2, 3))
	// Expected: 7

	tree := ast.NewBinaryExpression(
		ast.BinaryOperationAddition,
		ast.NewIntegerLiteral(1),
		ast.NewBinaryExpression(
			ast.BinaryOperationMultiplication,
			ast.NewIntegerLiteral(2),
			ast.NewIntegerLiteral(3),
		),
	)

	compiler := NewCompiler()
	err := compiler.Compile(tree)
	if err != nil {
		t.Fatalf("Compiler error: %v", err)
	}

	vm := NewVM(compiler.MainFunction(), nil)
	err = vm.Run()
	if err != nil {
		t.Fatalf("VM error: %v", err)
	}

	result := vm.StackTop()

	fmt.Printf("Result: %v (Type: %T)\n", result, result)

	if result != int64(7) {
		t.Errorf("Expected 7, got %v", result)
	}
}

func TestFloatMath(t *testing.T) {
	// 10.5 / 2
	tree := ast.NewBinaryExpression(
		ast.BinaryOperationDivision,
		ast.NewFloatLiteral(10.5),
		ast.NewIntegerLiteral(2),
	)

	compiler := NewCompiler()
	err := compiler.Compile(tree)
	if err != nil {
		t.Fatalf("Compiler error: %v", err)
	}

	vm := NewVM(compiler.MainFunction(), nil)
	err = vm.Run()
	if err != nil {
		t.Fatalf("VM error: %v", err)
	}

	result := vm.StackTop()
	if result != 5.25 {
		t.Errorf("Expected 5.25, got %v", result)
	}
}
