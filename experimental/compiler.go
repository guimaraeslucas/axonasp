package experimental

import (
	"encoding/binary"
	"fmt"

	"strings"

	"g3pix.com.br/axonasp/vbscript/ast"
)

// Scope represents a compilation scope (Global or Function)
type Scope struct {
	Name   string
	Locals []string
}

// Compiler maintains the state of the compilation process
type Compiler struct {
	instructions    []byte
	constants       []Value
	globalNames     []string
	globalIndex     map[string]int
	scopes          []*Scope
	procedures      map[string]bool
	declaredGlobals map[string]bool
	fallbackCount   int
}

type RequiresASTFallbackError struct {
	Reason string
}

func (e *RequiresASTFallbackError) Error() string {
	if e == nil || e.Reason == "" {
		return "vm compilation requires AST fallback"
	}
	return e.Reason
}

// NewCompiler creates a new Compiler instance
func NewCompiler() *Compiler {
	return &Compiler{
		instructions:    []byte{},
		constants:       []Value{},
		globalNames:     []string{},
		globalIndex:     make(map[string]int),
		scopes:          []*Scope{{Name: "global"}},
		procedures:      make(map[string]bool),
		declaredGlobals: make(map[string]bool),
	}
}

// Bytecode returns the compiled bytecode chunk
func (c *Compiler) Bytecode() *Bytecode {
	return &Bytecode{
		Instructions: c.instructions,
		Constants:    c.constants,
		GlobalNames:  c.globalNames,
	}
}

// MainFunction returns a Function object representing the top-level script
func (c *Compiler) MainFunction() *Function {
	return &Function{
		Name:           "main",
		Bytecode:       c.Bytecode(),
		ParameterCount: 0,
		LocalCount:     len(c.scopes[0].Locals),
	}
}

func (c *Compiler) currentScope() *Scope {
	return c.scopes[len(c.scopes)-1]
}

func (c *Compiler) enterScope(name string) {
	c.scopes = append(c.scopes, &Scope{Name: name})
}

func (c *Compiler) leaveScope() *Scope {
	s := c.currentScope()
	c.scopes = c.scopes[:len(c.scopes)-1]
	return s
}

func (c *Compiler) resolveVariable(name string) (int, bool, bool) {
	name = strings.ToLower(name)

	// Check current scope for locals
	s := c.currentScope()
	for i, local := range s.Locals {
		if local == name {
			return i, true, true
		}
	}

	// Globals
	globalIdx := c.resolveGlobalIndex(name)
	return globalIdx, false, true
}

func (c *Compiler) defineLocal(name string) int {
	name = strings.ToLower(name)
	s := c.currentScope()
	for i, local := range s.Locals {
		if local == name {
			return i
		}
	}
	s.Locals = append(s.Locals, name)
	return len(s.Locals) - 1
}

func (c *Compiler) resolveGlobalIndex(name string) int {
	name = strings.ToLower(name)
	if idx, ok := c.globalIndex[name]; ok {
		return idx
	}
	idx := len(c.globalNames)
	c.globalNames = append(c.globalNames, name)
	c.globalIndex[name] = idx
	if idx > 65535 {
		panic("too many globals")
	}
	return idx
}

func (c *Compiler) isGlobalScope() bool {
	return len(c.scopes) == 1
}

// Compile traverses the AST and generates bytecode
func (c *Compiler) Compile(node ast.Node) error {
	switch n := node.(type) {
	// --- Statements ---
	case *ast.Program:
		c.collectProcedureDeclarations(n.Body)
		for _, stmt := range n.Body {
			if err := c.Compile(stmt); err != nil {
				return err
			}
		}

	case *ast.StatementList:
		c.collectProcedureDeclarations(n.Statements)
		for _, stmt := range n.Statements {
			if err := c.Compile(stmt); err != nil {
				return err
			}
		}

	case *ast.AssignmentStatement:
		if mem, ok := n.Left.(*ast.MemberExpression); ok {
			_ = mem
			if err := c.emitFallback(n); err != nil {
				return err
			}
			return nil
		}

		if idxExpr, ok := n.Left.(*ast.IndexOrCallExpression); ok {
			_ = idxExpr
			if err := c.emitFallback(n); err != nil {
				return err
			}
			return nil
		}

		id, ok := n.Left.(*ast.Identifier)
		if !ok {
			return fmt.Errorf("unsupported assignment target: %T", n.Left)
		}

		if c.tryEmitIncrement(id, n.Right) {
			return nil
		}

		// Right side (Value)
		if err := c.Compile(n.Right); err != nil {
			return err
		}

		idx, isLocal, _ := c.resolveVariable(id.Name)
		if isLocal {
			c.emit(OP_SET_LOCAL, idx)
		} else {
			c.declaredGlobals[strings.ToLower(id.Name)] = true
			c.emit(OP_SET_GLOBAL_FAST, idx)
		}

	case *ast.VariablesDeclaration:
		for _, v := range n.Variables {
			if c.isGlobalScope() {
				idx := c.resolveGlobalIndex(v.Identifier.Name)
				c.declaredGlobals[strings.ToLower(v.Identifier.Name)] = true
				c.emit(OP_EMPTY)
				c.emit(OP_SET_GLOBAL_FAST, idx)
				continue
			}
			c.defineLocal(v.Identifier.Name)
		}

	case *ast.ConstsDeclaration:
		for _, decl := range n.Declarations {
			if decl == nil || decl.Identifier == nil {
				continue
			}
			if decl.Init != nil {
				if err := c.Compile(decl.Init); err != nil {
					return err
				}
			} else {
				c.emit(OP_EMPTY)
			}
			idx := c.resolveGlobalIndex(decl.Identifier.Name)
			c.emit(OP_SET_GLOBAL_FAST, idx)
		}

	case *ast.IfStatement:
		if err := c.Compile(n.Test); err != nil {
			return err
		}
		jumpIfFalsePos := c.emit(OP_JUMP_IF_FALSE, 0xFFFF)
		if err := c.Compile(n.Consequent); err != nil {
			return err
		}
		jumpToEndPos := c.emit(OP_JUMP, 0xFFFF)
		c.patchJump(jumpIfFalsePos)
		if n.Alternate != nil {
			if err := c.Compile(n.Alternate); err != nil {
				return err
			}
		}
		c.patchJump(jumpToEndPos)

	case *ast.CallSubStatement:
		if err := c.emitFallback(n); err != nil {
			return err
		}

	case *ast.CallStatement:
		if err := c.emitFallback(n); err != nil {
			return err
		}

	case *ast.SubDeclaration:
		if n.Identifier == nil {
			return nil
		}
		params := make([]string, 0, len(n.Parameters))
		for _, p := range n.Parameters {
			if p != nil && p.Identifier != nil {
				params = append(params, p.Identifier.Name)
			}
		}

		bodyStatements := extractStatementList(n.Body)
		fn, err := c.CompileFunction(n.Identifier.Name, params, bodyStatements, true)
		if err != nil {
			return err
		}
		idx := c.addConstant(fn)
		c.emit(OP_CONSTANT, idx)
		nameIdx := c.resolveGlobalIndex(n.Identifier.Name)
		c.emit(OP_SET_GLOBAL_FAST, nameIdx)

	case *ast.FunctionDeclaration:
		if n.Identifier == nil {
			return nil
		}
		params := make([]string, 0, len(n.Parameters))
		for _, p := range n.Parameters {
			if p != nil && p.Identifier != nil {
				params = append(params, p.Identifier.Name)
			}
		}

		bodyStatements := extractStatementList(n.Body)
		fn, err := c.CompileFunction(n.Identifier.Name, params, bodyStatements, false)
		if err != nil {
			return err
		}
		idx := c.addConstant(fn)
		c.emit(OP_CONSTANT, idx)
		nameIdx := c.resolveGlobalIndex(n.Identifier.Name)
		c.emit(OP_SET_GLOBAL_FAST, nameIdx)

	case *ast.ClassDeclaration:
		if n.Identifier == nil {
			return nil
		}

		compiledClass := &CompiledClass{
			Name:    n.Identifier.Name,
			Methods: make(map[string]*Function),
		}

		// Compile methods and collect variables
		for _, member := range n.Members {
			switch m := member.(type) {
			case *ast.FunctionDeclaration:
				params := make([]string, 0, len(m.Parameters))
				for _, p := range m.Parameters {
					if p != nil && p.Identifier != nil {
						params = append(params, p.Identifier.Name)
					}
				}
				bodyStmts := extractStatementList(m.Body)
				fn, err := c.CompileFunction(m.Identifier.Name, params, bodyStmts, false)
				if err != nil {
					return err
				}
				compiledClass.Methods[strings.ToLower(m.Identifier.Name)] = fn

			case *ast.SubDeclaration:
				params := make([]string, 0, len(m.Parameters))
				for _, p := range m.Parameters {
					if p != nil && p.Identifier != nil {
						params = append(params, p.Identifier.Name)
					}
				}
				bodyStmts := extractStatementList(m.Body)
				fn, err := c.CompileFunction(m.Identifier.Name, params, bodyStmts, true)
				if err != nil {
					return err
				}
				compiledClass.Methods[strings.ToLower(m.Identifier.Name)] = fn

			case *ast.VariablesDeclaration:
				for _, v := range m.Variables {
					compiledClass.Variables = append(compiledClass.Variables, v.Identifier.Name)
				}
			case *ast.FieldsDeclaration:
				for _, f := range m.Fields {
					compiledClass.Variables = append(compiledClass.Variables, f.Identifier.Name)
				}
			}
		}

		idx := c.addConstant(compiledClass)
		c.emit(OP_CONSTANT, idx)
		nameIdx := c.resolveGlobalIndex(n.Identifier.Name)
		c.emit(OP_SET_GLOBAL_FAST, nameIdx)

	// --- Expressions ---
	case *ast.Identifier:
		name := strings.ToLower(n.Name)
		idx, isLocal, _ := c.resolveVariable(name)
		if isLocal {
			c.emit(OP_GET_LOCAL, idx)
		} else {
			c.emit(OP_GET_GLOBAL_FAST, idx)
		}

	case *ast.IndexOrCallExpression:
		if err := c.emitFallback(n); err != nil {
			return err
		}

	case *ast.MemberExpression:
		if err := c.emitFallback(n); err != nil {
			return err
		}

	case *ast.IntegerLiteral:
		c.emitConstant(n.Value)

	case *ast.EmptyLiteral:
		c.emit(OP_EMPTY)

	case *ast.NothingLiteral:
		c.emit(OP_NOTHING)

	case *ast.FloatLiteral:
		c.emitConstant(n.Value)

	case *ast.StringLiteral:
		c.emitConstant(n.Value)

	case *ast.BooleanLiteral:
		if n.Value {
			c.emit(OP_TRUE)
		} else {
			c.emit(OP_FALSE)
		}

	case *ast.NewExpression:
		if n.Argument == nil {
			return fmt.Errorf("invalid New expression")
		}
		switch arg := n.Argument.(type) {
		case *ast.Identifier:
			idx := c.addConstant(arg.Name)
			c.emit(OP_NEW, idx)
		case *ast.StringLiteral:
			idx := c.addConstant(arg.Value)
			c.emit(OP_NEW, idx)
		default:
			return fmt.Errorf("unsupported New expression argument: %T", n.Argument)
		}

	case *ast.BinaryExpression:
		if err := c.Compile(n.Left); err != nil {
			return err
		}
		if err := c.Compile(n.Right); err != nil {
			return err
		}

		switch n.Operation {
		case ast.BinaryOperationExponentiation:
			c.emit(OP_POW)
		case ast.BinaryOperationAddition:
			c.emit(OP_ADD)
		case ast.BinaryOperationSubtraction:
			c.emit(OP_SUB)
		case ast.BinaryOperationMultiplication:
			c.emit(OP_MUL)
		case ast.BinaryOperationDivision:
			c.emit(OP_DIV)
		case ast.BinaryOperationIntDivision:
			c.emit(OP_IDIV)
		case ast.BinaryOperationMod:
			c.emit(OP_MOD)
		case ast.BinaryOperationConcatenation:
			c.emit(OP_CONCAT)
		case ast.BinaryOperationEqual:
			c.emit(OP_EQUAL)
		case ast.BinaryOperationNotEqual:
			c.emit(OP_NOT_EQUAL)
		case ast.BinaryOperationLess:
			c.emit(OP_LESS)
		case ast.BinaryOperationLessOrEqual:
			c.emit(OP_LESS_EQUAL)
		case ast.BinaryOperationGreater:
			c.emit(OP_GREATER)
		case ast.BinaryOperationGreaterOrEqual:
			c.emit(OP_GREATER_EQUAL)
		case ast.BinaryOperationIs:
			c.emit(OP_IS)
		case ast.BinaryOperationAnd:
			c.emit(OP_AND)
		case ast.BinaryOperationOr:
			c.emit(OP_OR)
		case ast.BinaryOperationXor:
			c.emit(OP_XOR)
		case ast.BinaryOperationEqv:
			c.emit(OP_EQV)
		case ast.BinaryOperationImp:
			c.emit(OP_IMP)
		default:
			return fmt.Errorf("unknown binary operator: %d", n.Operation)
		}

	case *ast.UnaryExpression:
		if err := c.Compile(n.Argument); err != nil {
			return err
		}
		switch n.Operation {
		case ast.UnaryOperationMinus:
			c.emit(OP_NEG)
		case ast.UnaryOperationNot:
			c.emit(OP_NOT)
		default:
			return fmt.Errorf("unknown unary operator: %d", n.Operation)
		}

	default:
		// Fallback to AST walker for unimplemented nodes
		if err := c.emitFallback(node); err != nil {
			return err
		}
	}

	return nil
}

func (c *Compiler) emitFallback(node ast.Node) error {
	_ = node
	c.fallbackCount++
	return &RequiresASTFallbackError{Reason: "vm requires ast fallback for unsupported node"}
}

func (c *Compiler) tryEmitIncrement(left *ast.Identifier, right ast.Expression) bool {
	bin, ok := right.(*ast.BinaryExpression)
	if !ok {
		return false
	}
	if bin.Operation != ast.BinaryOperationAddition {
		return false
	}
	leftIdent, ok := bin.Left.(*ast.Identifier)
	if !ok {
		return false
	}
	if strings.ToLower(leftIdent.Name) != strings.ToLower(left.Name) {
		return false
	}
	lit, ok := bin.Right.(*ast.IntegerLiteral)
	if !ok || lit.Value != 1 {
		return false
	}

	idx, isLocal, _ := c.resolveVariable(left.Name)
	if isLocal {
		c.emit(OP_INC_LOCAL, idx)
	} else {
		c.emit(OP_INC_GLOBAL_FAST, idx)
	}
	return true
}

// emit adds an instruction to the bytecode
func (c *Compiler) emit(op Opcode, operands ...int) int {
	pos := len(c.instructions)
	c.instructions = append(c.instructions, byte(op))

	def, ok := Lookup(byte(op))
	if !ok {
		panic(fmt.Sprintf("undefined opcode: %d", op))
	}

	if len(operands) != len(def.OperandWidths) {
		panic(fmt.Sprintf("opcode %s expected %d operands, got %d", def.Name, len(def.OperandWidths), len(operands)))
	}

	for i, width := range def.OperandWidths {
		operand := operands[i]
		switch width {
		case 1:
			c.instructions = append(c.instructions, byte(operand))
		case 2:
			// BigEndian for consistency
			b := make([]byte, 2)
			binary.BigEndian.PutUint16(b, uint16(operand))
			c.instructions = append(c.instructions, b...)
		}
	}

	return pos
}

// patchJump calculates the offset from the jump instruction to the current position
// and overwrites the placeholder operand.
func (c *Compiler) patchJump(opPos int) {
	// Opcode is at opPos. Operand (2 bytes) starts at opPos + 1.
	// Target is len(c.instructions).
	// Jump is relative to the IP AFTER reading operands.
	// So offset = target - (opPos + 1 + 2) = target - opPos - 3

	target := len(c.instructions)
	offset := target - opPos - 3

	if offset > 65535 {
		panic("jump offset too large")
	}

	binary.BigEndian.PutUint16(c.instructions[opPos+1:], uint16(offset))
}

// addConstant adds a value to the constant pool and returns its index.
// It tries to find an existing identical constant first.
func (c *Compiler) addConstant(v Value) int {
	// Simple linear search for now (Phase 2)
	for i, existing := range c.constants {
		if existing == v {
			return i
		}
	}

	c.constants = append(c.constants, v)
	idx := len(c.constants) - 1
	if idx > 65535 {
		panic("too many constants")
	}
	return idx
}

// emitConstant adds a value to the constant pool and emits an OP_CONSTANT instruction
func (c *Compiler) emitConstant(v Value) {
	idx := c.addConstant(v)
	c.emit(OP_CONSTANT, idx)
}

func (c *Compiler) collectProcedureDeclarations(stmts []ast.Statement) {
	for _, stmt := range stmts {
		switch decl := stmt.(type) {
		case *ast.SubDeclaration:
			if decl.Identifier != nil {
				c.procedures[strings.ToLower(decl.Identifier.Name)] = true
			}
		case *ast.FunctionDeclaration:
			if decl.Identifier != nil {
				c.procedures[strings.ToLower(decl.Identifier.Name)] = true
			}
		}
	}
}

func (c *Compiler) isKnownVariable(nameLower string) bool {
	if c.currentScope() != nil {
		for _, local := range c.currentScope().Locals {
			if local == nameLower {
				return true
			}
		}
	}
	return c.declaredGlobals[nameLower]
}

func extractStatementList(stmt ast.Statement) []ast.Statement {
	if stmt == nil {
		return nil
	}
	if list, ok := stmt.(*ast.StatementList); ok {
		return list.Statements
	}
	return []ast.Statement{stmt}
}

func formatMemberName(expr *ast.MemberExpression) (string, bool) {
	if expr == nil || expr.Property == nil {
		return "", false
	}
	objIdent, ok := expr.Object.(*ast.Identifier)
	if !ok || objIdent == nil {
		return "", false
	}
	return strings.ToLower(objIdent.Name + "." + expr.Property.Name), true
}

// CompileFunction compiles a Sub or Function declaration
func (c *Compiler) CompileFunction(name string, params []string, body []ast.Statement, isSub bool) (*Function, error) {
	c.enterScope(name)

	// Define parameters as locals
	for _, p := range params {
		c.defineLocal(p)
	}

	// Define function name as a local for the return value
	if !isSub {
		c.defineLocal(name)
	}

	// Save current instructions/constants and start fresh for the function
	oldInstructions := c.instructions
	c.instructions = []byte{}

	for _, stmt := range body {
		if err := c.Compile(stmt); err != nil {
			return nil, err
		}
	}

	// Add implicit return
	if isSub {
		c.emit(OP_RETURN)
	} else {
		// Functions in VBScript return the value assigned to their name.
		// We need to resolve the function name as a local variable.
		idx, isLocal, _ := c.resolveVariable(name)
		if isLocal {
			c.emit(OP_GET_LOCAL, idx)
		} else {
			c.emit(OP_EMPTY)
		}
		c.emit(OP_RETURN_VALUE)
	}

	funcInstructions := c.instructions
	c.instructions = oldInstructions

	scope := c.leaveScope()

	return &Function{
		Name:           name,
		Bytecode:       &Bytecode{Instructions: funcInstructions, Constants: c.constants},
		ParameterCount: len(params),
		LocalCount:     len(scope.Locals),
	}, nil
}
