package server

import (
	"fmt"
	"go-asp/asp"
	"math"
	"net/http"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/guimaraeslucas/vbscript-go/ast"
)

// LoopExitError represents a loop exit statement (Exit For, Exit Do, etc)
type LoopExitError struct {
	LoopType string // "for", "do", "while", "select"
}

func (e *LoopExitError) Error() string {
	return fmt.Sprintf("Exit %s", e.LoopType)
}

// ExecutionContext holds all runtime state for ASP execution
type ExecutionContext struct {
	// ASP core objects
	Request     *asp.RequestObject
	Response    *asp.ResponseObject
	Server      *asp.ServerObject
	Session     map[string]interface{}
	Application map[string]interface{}

	// Variable storage (case-insensitive keys)
	variables map[string]interface{}

	// HTTP context
	httpWriter  http.ResponseWriter
	httpRequest *http.Request

	// Execution state
	startTime time.Time
	timeout   time.Duration

	// Library instances
	libraries map[string]interface{}

	// Mutex for thread safety
	mu sync.RWMutex
}

// NewExecutionContext creates a new execution context
func NewExecutionContext(w http.ResponseWriter, r *http.Request, sessionID string, timeout time.Duration) *ExecutionContext {
	return &ExecutionContext{
		Request:     asp.NewRequestObject(),
		Response:    asp.NewResponseObject(),
		Server:      asp.NewServerObject(),
		Session:     make(map[string]interface{}),
		Application: make(map[string]interface{}),
		variables:   make(map[string]interface{}),
		libraries:   make(map[string]interface{}),
		httpWriter:  w,
		httpRequest: r,
		startTime:   time.Now(),
		timeout:     timeout,
	}
}

// SetVariable sets a variable in the execution context (case-insensitive)
func (ec *ExecutionContext) SetVariable(name string, value interface{}) {
	ec.mu.Lock()
	defer ec.mu.Unlock()
	ec.variables[strings.ToLower(name)] = value
}

// GetVariable gets a variable from the execution context (case-insensitive)
func (ec *ExecutionContext) GetVariable(name string) (interface{}, bool) {
	ec.mu.RLock()
	defer ec.mu.RUnlock()
	val, exists := ec.variables[strings.ToLower(name)]
	return val, exists
}

// CheckTimeout checks if execution has exceeded timeout
func (ec *ExecutionContext) CheckTimeout() error {
	if time.Since(ec.startTime) > ec.timeout {
		return fmt.Errorf("execution timeout exceeded (%v)", ec.timeout)
	}
	return nil
}

// ASPExecutor handles execution of ASP code with VBScript programs
type ASPExecutor struct {
	config  *ASPProcessorConfig
	context *ExecutionContext
}

// NewASPExecutor creates a new ASP executor
func NewASPExecutor(config *ASPProcessorConfig) *ASPExecutor {
	if config == nil {
		config = &ASPProcessorConfig{
			RootDir:       "./www",
			ScriptTimeout: 30,
		}
	}

	return &ASPExecutor{
		config: config,
	}
}

// Execute processes ASP code and returns rendered output
func (ae *ASPExecutor) Execute(fileContent string, w http.ResponseWriter, r *http.Request, sessionID string) error {
	// Create execution context
	timeout := time.Duration(ae.config.ScriptTimeout) * time.Second
	ae.context = NewExecutionContext(w, r, sessionID, timeout)

	// Configure Server object with context
	ae.context.Server.SetProperty("_rootDir", ae.config.RootDir)
	ae.context.Server.SetProperty("_httpRequest", r)

	// Populate Request object
	populateRequestData(ae.context.Request, r)

	// Parse ASP code
	parser := asp.NewASPParser(fileContent)
	result, err := parser.Parse()
	if err != nil {
		return fmt.Errorf("failed to parse ASP code: %w", err)
	}

	// Check for parse errors
	if len(result.Errors) > 0 {
		return fmt.Errorf("ASP parse error: %v", result.Errors[0])
	}

	// Execute blocks in order with timeout protection
	done := make(chan error, 1)

	go func() {
		defer func() {
			if rec := recover(); rec != nil {
				done <- fmt.Errorf("runtime panic: %v", rec)
			}
		}()

		err := ae.executeBlocks(result)
		done <- err
	}()

	// Wait for execution or timeout
	select {
	case err := <-done:
		if err != nil {
			return err
		}
	case <-time.After(timeout):
		return fmt.Errorf("script execution timeout (>%d seconds)", ae.config.ScriptTimeout)
	}

	// Write response to HTTP ResponseWriter
	buffer := ae.context.Response.GetBuffer()
	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	_, err = w.Write([]byte(buffer))
	if err != nil {
		return fmt.Errorf("failed to write response: %w", err)
	}

	return nil
}

// executeBlocks executes all blocks in order (HTML and ASP)
func (ae *ASPExecutor) executeBlocks(result *asp.ASPParserResult) error {
	for i, block := range result.Blocks {
		// Check timeout periodically
		if i%100 == 0 {
			if err := ae.context.CheckTimeout(); err != nil {
				return err
			}
		}

		switch block.Type {
		case "html":
			// Write HTML content directly
			ae.context.Response.CallMethod("Write", block.Content)

		case "asp":
			// Execute VBScript block if parsed
			if program, exists := result.VBPrograms[i]; exists && program != nil {
				if err := ae.executeVBProgram(program); err != nil {
					return err
				}
			}
		}
	}

	return nil
}

// executeVBProgram executes a VBScript AST program
func (ae *ASPExecutor) executeVBProgram(program *ast.Program) error {
	if program == nil {
		return nil
	}

	// Create a visitor to traverse the AST
	v := NewASPVisitor(ae.context, ae)

	// Visit all statements in the program
	if program.Body != nil {
		for _, stmt := range program.Body {
			if stmt == nil {
				continue
			}

			// Check timeout
			if err := ae.context.CheckTimeout(); err != nil {
				return err
			}

			// Execute statement
			if err := v.VisitStatement(stmt); err != nil {
				return err
			}
		}
	}

	return nil
}

// CreateObject creates an ASP COM object (like Server.CreateObject)
func (ae *ASPExecutor) CreateObject(objType string) (interface{}, error) {
	objType = strings.ToUpper(objType)

	switch objType {
	case "G3JSON":
		return NewJSONLibrary(ae.context), nil
	case "G3FILES":
		return NewFileSystemLibrary(ae.context), nil
	case "G3HTTP":
		return NewHTTPLibrary(ae.context), nil
	case "G3TEMPLATE":
		return NewTemplateLibrary(ae.context), nil
	case "G3MAIL":
		return NewMailLibrary(ae.context), nil
	case "G3CRYPTO":
		return NewCryptoLibrary(ae.context), nil
	case "MSXML2.SERVERXMLHTTP":
		return NewServerXMLHTTP(ae.context), nil
	case "MSXML2.DOMDOCUMENT":
		return NewDOMDocument(ae.context), nil
	case "ADODB.CONNECTION":
		return NewADOConnection(ae.context), nil
	case "ADODB.RECORDSET":
		return NewADORecordset(ae.context), nil
	case "ADODB.STREAM":
		return NewADOStream(ae.context), nil
	default:
		return nil, fmt.Errorf("unsupported object type: %s", objType)
	}
}

// ASPVisitor traverses and executes the VBScript AST
type ASPVisitor struct {
	context  *ExecutionContext
	executor *ASPExecutor
	depth    int
}

// NewASPVisitor creates a new ASP visitor for AST traversal
func NewASPVisitor(ctx *ExecutionContext, executor *ASPExecutor) *ASPVisitor {
	return &ASPVisitor{
		context:  ctx,
		executor: executor,
		depth:    0,
	}
}

// VisitStatement executes a single statement from the AST
func (v *ASPVisitor) VisitStatement(node ast.Statement) error {
	if node == nil {
		return nil
	}

	v.depth++
	if v.depth > 1000 {
		return fmt.Errorf("maximum call depth exceeded")
	}
	defer func() { v.depth-- }()

	switch stmt := node.(type) {
	case *ast.AssignmentStatement:
		return v.visitAssignment(stmt)

	case *ast.CallStatement:
		_, err := v.visitExpression(stmt.Callee)
		return err

	case *ast.ReDimStatement:
		return v.visitReDim(stmt)

	case *ast.IfStatement:
		return v.visitIf(stmt)

	case *ast.ForStatement:
		return v.visitFor(stmt)

	case *ast.ForEachStatement:
		return v.visitForEach(stmt)

	case *ast.DoStatement:
		return v.visitDo(stmt)

	case *ast.WhileStatement:
		return v.visitWhile(stmt)

	case *ast.SelectStatement:
		return v.visitSelect(stmt)

	case *ast.SubDeclaration:
		return v.visitSubDeclaration(stmt)

	case *ast.FunctionDeclaration:
		return v.visitFunctionDeclaration(stmt)

	case *ast.ClassDeclaration:
		return v.visitClassDeclaration(stmt)

	case *ast.OnErrorResumeNextStatement:
		// Error handling - continue on error
		return nil

	case *ast.OnErrorGoTo0Statement:
		// Error handling - reset error
		return nil

	default:
		// Try to evaluate as expression for side effects
		if expr, ok := node.(ast.Expression); ok {
			_, err := v.visitExpression(expr)
			return err
		}
	}

	return nil
}

// visitAssignment handles variable assignment
func (v *ASPVisitor) visitAssignment(stmt *ast.AssignmentStatement) error {
	if stmt == nil || stmt.Right == nil {
		return nil
	}

	// Evaluate right-hand side
	value, err := v.visitExpression(stmt.Right)
	if err != nil {
		return err
	}

	// Get variable name from left side
	if ident, ok := stmt.Left.(*ast.Identifier); ok {
		v.context.SetVariable(ident.Name, value)
	}

	return nil
}

// visitReDim handles ReDim statements
func (v *ASPVisitor) visitReDim(stmt *ast.ReDimStatement) error {
	if stmt == nil || stmt.ReDims == nil {
		return nil
	}

	for _, redim := range stmt.ReDims {
		if redim == nil || redim.Identifier == nil {
			continue
		}
		varName := redim.Identifier.Name
		// Initialize array - in full implementation, respect dimensions
		v.context.SetVariable(varName, make([]interface{}, 0))
	}

	return nil
}

// visitIf handles if-else statements
func (v *ASPVisitor) visitIf(stmt *ast.IfStatement) error {
	if stmt == nil || stmt.Test == nil {
		return nil
	}

	condition, err := v.visitExpression(stmt.Test)
	if err != nil {
		return err
	}

	// Convert condition to boolean
	if isTruthy(condition) {
		// Execute consequent block
		if stmt.Consequent != nil {
			if err := v.VisitStatement(stmt.Consequent); err != nil {
				return err
			}
		}
	} else {
		// Execute alternate block
		if stmt.Alternate != nil {
			if err := v.VisitStatement(stmt.Alternate); err != nil {
				return err
			}
		}
	}

	return nil
}

// visitFor handles for loops
func (v *ASPVisitor) visitFor(stmt *ast.ForStatement) error {
	if stmt == nil || stmt.From == nil || stmt.To == nil {
		return nil
	}

	// Get variable name
	var varName string
	if stmt.Identifier != nil {
		varName = stmt.Identifier.Name
	}
	if varName == "" {
		return nil
	}

	// Evaluate From and To
	from, err := v.visitExpression(stmt.From)
	if err != nil {
		return err
	}

	to, err := v.visitExpression(stmt.To)
	if err != nil {
		return err
	}

	// Evaluate Step (default 1)
	step := 1.0
	if stmt.Step != nil {
		stepVal, err := v.visitExpression(stmt.Step)
		if err != nil {
			return err
		}
		step = toFloat(stepVal)
	}

	// Loop
	current := toFloat(from)
	end := toFloat(to)

	if step > 0 {
		for current <= end {
			v.context.SetVariable(varName, current)

			// Execute body
			if stmt.Body != nil {
				for _, s := range stmt.Body {
					if err := v.VisitStatement(s); err != nil {
						return err
					}
				}
			}

			current += step
		}
	} else if step < 0 {
		for current >= end {
			v.context.SetVariable(varName, current)

			// Execute body
			if stmt.Body != nil {
				for _, s := range stmt.Body {
					if err := v.VisitStatement(s); err != nil {
						return err
					}
				}
			}

			current += step
		}
	}

	return nil
}

// visitForEach handles for-each loops
func (v *ASPVisitor) visitForEach(stmt *ast.ForEachStatement) error {
	if stmt == nil || stmt.Identifier == nil {
		return nil
	}

	// Evaluate the collection expression
	collection, err := v.visitExpression(stmt.In)
	if err != nil {
		return err
	}

	// Handle different collection types
	switch col := collection.(type) {
	case []interface{}:
		// Iterate over array
		for _, item := range col {
			// Set loop variable
			v.context.SetVariable(stmt.Identifier.Name, item)

			// Execute loop body
			for _, body := range stmt.Body {
				if err := v.VisitStatement(body); err != nil {
					// Handle Exit For
					if _, ok := err.(*LoopExitError); ok && err.(*LoopExitError).LoopType == "for" {
						break
					}
					return err
				}
			}
		}
	case map[string]interface{}:
		// Iterate over map (VBScript dictionary)
		for key := range col {
			// Set loop variable to key
			v.context.SetVariable(stmt.Identifier.Name, key)

			// Execute loop body
			for _, body := range stmt.Body {
				if err := v.VisitStatement(body); err != nil {
					// Handle Exit For
					if _, ok := err.(*LoopExitError); ok && err.(*LoopExitError).LoopType == "for" {
						break
					}
					return err
				}
			}
		}
	}

	return nil
}

// visitDo handles do-while loops
func (v *ASPVisitor) visitDo(stmt *ast.DoStatement) error {
	if stmt == nil {
		return nil
	}

	for {
		// Check pre-test condition if needed
		if stmt.TestType == ast.ConditionTestTypePreTest {
			condition, err := v.visitExpression(stmt.Condition)
			if err != nil {
				return err
			}

			// Handle loop type (While vs Until)
			shouldContinue := isTruthy(condition)
			if stmt.LoopType == ast.LoopTypeUntil {
				shouldContinue = !shouldContinue
			}

			if !shouldContinue {
				break
			}
		}

		// Execute loop body
		for _, body := range stmt.Body {
			if err := v.VisitStatement(body); err != nil {
				// Handle Exit Do
				if _, ok := err.(*LoopExitError); ok && err.(*LoopExitError).LoopType == "do" {
					return nil
				}
				return err
			}
		}

		// Check post-test condition if needed
		if stmt.TestType == ast.ConditionTestTypePostTest {
			condition, err := v.visitExpression(stmt.Condition)
			if err != nil {
				return err
			}

			// Handle loop type (While vs Until)
			shouldContinue := isTruthy(condition)
			if stmt.LoopType == ast.LoopTypeUntil {
				shouldContinue = !shouldContinue
			}

			if !shouldContinue {
				break
			}
		}
	}

	return nil
}

// visitWhile handles while loops
func (v *ASPVisitor) visitWhile(stmt *ast.WhileStatement) error {
	if stmt == nil {
		return nil
	}

	for {
		condition, err := v.visitExpression(stmt.Condition)
		if err != nil {
			return err
		}

		if !isTruthy(condition) {
			break
		}

		// Execute loop body
		for _, body := range stmt.Body {
			if err := v.VisitStatement(body); err != nil {
				// Handle Exit While
				if _, ok := err.(*LoopExitError); ok && err.(*LoopExitError).LoopType == "while" {
					return nil
				}
				return err
			}
		}
	}

	return nil
}

// visitSelect handles select-case statements
func (v *ASPVisitor) visitSelect(stmt *ast.SelectStatement) error {
	if stmt == nil {
		return nil
	}

	// Evaluate select expression
	selectValue, err := v.visitExpression(stmt.Condition)
	if err != nil {
		return err
	}

	// Check each case
	for _, caseStmt := range stmt.Cases {
		if caseStmt == nil {
			continue
		}

		// Check if case matches
		matched := false
		if len(caseStmt.Values) == 0 {
			// Case Else
			matched = true
		} else {
			for _, caseValue := range caseStmt.Values {
				val, err := v.visitExpression(caseValue)
				if err != nil {
					return err
				}

				// Compare case value with select value
				if compareEqual(selectValue, val) {
					matched = true
					break
				}
			}
		}

		// Execute case body if matched
		if matched {
			for _, body := range caseStmt.Body {
				if err := v.VisitStatement(body); err != nil {
					// Handle Exit Select
					if _, ok := err.(*LoopExitError); ok && err.(*LoopExitError).LoopType == "select" {
						return nil
					}
					return err
				}
			}
			// Don't continue to next case (VBScript behavior)
			break
		}
	}

	return nil
}

// visitSubDeclaration handles sub declarations
func (v *ASPVisitor) visitSubDeclaration(stmt *ast.SubDeclaration) error {
	if stmt == nil || stmt.Identifier == nil {
		return nil
	}

	// Store sub in context for later calls
	v.context.SetVariable(stmt.Identifier.Name, stmt)
	return nil
}

// visitFunctionDeclaration handles function declarations
func (v *ASPVisitor) visitFunctionDeclaration(stmt *ast.FunctionDeclaration) error {
	if stmt == nil || stmt.Identifier == nil {
		return nil
	}

	// Store function in context for later calls
	v.context.SetVariable(stmt.Identifier.Name, stmt)
	return nil
}

// visitClassDeclaration handles class declarations
func (v *ASPVisitor) visitClassDeclaration(stmt *ast.ClassDeclaration) error {
	if stmt == nil || stmt.Identifier == nil {
		return nil
	}

	// Store class in context
	v.context.SetVariable(stmt.Identifier.Name, stmt)
	return nil
}

// visitExpression evaluates an expression and returns its value
func (v *ASPVisitor) visitExpression(expr ast.Expression) (interface{}, error) {
	if expr == nil {
		return nil, nil
	}

	switch e := expr.(type) {
	case *ast.Identifier:
		varName := e.Name
		if val, exists := v.context.GetVariable(varName); exists {
			return val, nil
		}
		// Undefined variable returns nil in VBScript
		return nil, nil

	case *ast.StringLiteral:
		return e.Value, nil

	case *ast.IntegerLiteral:
		return int(e.Value), nil

	case *ast.FloatLiteral:
		return e.Value, nil

	case *ast.BooleanLiteral:
		return e.Value, nil

	case *ast.BinaryExpression:
		return v.visitBinaryExpression(e)

	case *ast.UnaryExpression:
		return v.visitUnaryExpression(e)

	case *ast.IndexOrCallExpression:
		return v.visitIndexOrCall(e)

	case *ast.MemberExpression:
		return v.visitMemberExpression(e)

	default:
		return nil, nil
	}
}

// visitBinaryExpression evaluates binary operations
func (v *ASPVisitor) visitBinaryExpression(expr *ast.BinaryExpression) (interface{}, error) {
	if expr == nil {
		return nil, nil
	}

	left, err := v.visitExpression(expr.Left)
	if err != nil {
		return nil, err
	}

	// Handle short-circuit evaluation
	switch expr.Operation {
	case ast.BinaryOperationAnd:
		if !isTruthy(left) {
			return false, nil
		}
		right, err := v.visitExpression(expr.Right)
		if err != nil {
			return nil, err
		}
		return isTruthy(right), nil

	case ast.BinaryOperationOr:
		if isTruthy(left) {
			return true, nil
		}
		right, err := v.visitExpression(expr.Right)
		if err != nil {
			return nil, err
		}
		return isTruthy(right), nil
	}

	right, err := v.visitExpression(expr.Right)
	if err != nil {
		return nil, err
	}

	// Perform operation
	return performBinaryOperation(expr.Operation, left, right)
}

// visitUnaryExpression evaluates unary operations
func (v *ASPVisitor) visitUnaryExpression(expr *ast.UnaryExpression) (interface{}, error) {
	if expr == nil {
		return nil, nil
	}

	operand, err := v.visitExpression(expr.Argument)
	if err != nil {
		return nil, err
	}

	switch expr.Operation {
	case ast.UnaryOperationNot:
		return !isTruthy(operand), nil
	case ast.UnaryOperationMinus:
		return negateValue(operand), nil
	case ast.UnaryOperationPlus:
		return operand, nil
	default:
		return nil, fmt.Errorf("unknown unary operation: %v", expr.Operation)
	}
}

// visitIndexOrCall handles function calls and array indexing
func (v *ASPVisitor) visitIndexOrCall(expr *ast.IndexOrCallExpression) (interface{}, error) {
	if expr == nil {
		return nil, nil
	}

	// Evaluate base expression
	base, err := v.visitExpression(expr.Object)
	if err != nil {
		return nil, err
	}

	// Evaluate indexes (arguments)
	args := make([]interface{}, 0)
	if expr.Indexes != nil {
		for _, arg := range expr.Indexes {
			val, err := v.visitExpression(arg)
			if err != nil {
				return nil, err
			}
			args = append(args, val)
		}
	}

	// Handle method calls on built-in objects
	if obj, ok := base.(asp.ASPObject); ok {
		// Get method name from object expression
		if ident, ok := expr.Object.(*ast.Identifier); ok {
			return obj.CallMethod(ident.Name, args...)
		}
	}

	// Handle array access
	if arr, ok := base.([]interface{}); ok && len(args) > 0 {
		idx := toInt(args[0])
		if idx >= 0 && idx < len(arr) {
			return arr[idx], nil
		}
		return nil, nil
	}

	return nil, nil
}

// visitMemberExpression evaluates member access (obj.property)
func (v *ASPVisitor) visitMemberExpression(expr *ast.MemberExpression) (interface{}, error) {
	if expr == nil {
		return nil, nil
	}

	// Evaluate object
	obj, err := v.visitExpression(expr.Object)
	if err != nil {
		return nil, err
	}

	// Get property name
	propName := ""
	if expr.Property != nil {
		propName = expr.Property.Name
	}

	// Handle ASP built-in objects
	switch strings.ToLower(propName) {
	case "response":
		return v.context.Response, nil
	case "request":
		return v.context.Request, nil
	case "server":
		return v.context.Server, nil
	case "session":
		return v.context.Session, nil
	case "application":
		return v.context.Application, nil
	}

	// Handle generic property access
	if aspObj, ok := obj.(asp.ASPObject); ok {
		return aspObj.GetProperty(propName), nil
	}

	return nil, nil
}

// Helper functions

// isTruthy checks if a value is truthy in VBScript
func isTruthy(val interface{}) bool {
	if val == nil {
		return false
	}
	if b, ok := val.(bool); ok {
		return b
	}
	if i, ok := val.(int); ok {
		return i != 0
	}
	if i, ok := val.(float64); ok {
		return i != 0
	}
	if s, ok := val.(string); ok {
		return s != ""
	}
	return true
}

// toString converts a value to string
func toString(val interface{}) string {
	if val == nil {
		return ""
	}
	switch v := val.(type) {
	case string:
		return v
	case bool:
		if v {
			return "True"
		}
		return "False"
	case int:
		return fmt.Sprintf("%d", v)
	case float64:
		if v == float64(int(v)) {
			return fmt.Sprintf("%d", int(v))
		}
		return fmt.Sprintf("%g", v)
	default:
		return fmt.Sprintf("%v", v)
	}
}

// toInt converts a value to integer
func toInt(val interface{}) int {
	if val == nil {
		return 0
	}
	switch v := val.(type) {
	case int:
		return v
	case float64:
		return int(v)
	case string:
		if i, err := strconv.Atoi(v); err == nil {
			return i
		}
		return 0
	case bool:
		if v {
			return -1
		}
		return 0
	default:
		return 0
	}
}

// toFloat converts a value to float64
func toFloat(val interface{}) float64 {
	if val == nil {
		return 0
	}
	switch v := val.(type) {
	case int:
		return float64(v)
	case float64:
		return v
	case string:
		if f, err := strconv.ParseFloat(v, 64); err == nil {
			return f
		}
		return 0
	case bool:
		if v {
			return -1
		}
		return 0
	default:
		return 0
	}
}

// negateValue negates a value
func negateValue(val interface{}) interface{} {
	if val == nil {
		return 0
	}
	switch v := val.(type) {
	case int:
		return -v
	case float64:
		return -v
	default:
		return 0
	}
}

// performBinaryOperation performs a binary operation
func performBinaryOperation(op ast.BinaryOperation, left, right interface{}) (interface{}, error) {
	switch op {
	case ast.BinaryOperationAnd:
		return isTruthy(left) && isTruthy(right), nil
	case ast.BinaryOperationOr:
		return isTruthy(left) || isTruthy(right), nil
	case ast.BinaryOperationAddition:
		leftNum := toFloat(left)
		rightNum := toFloat(right)
		return leftNum + rightNum, nil
	case ast.BinaryOperationSubtraction:
		leftNum := toFloat(left)
		rightNum := toFloat(right)
		return leftNum - rightNum, nil
	case ast.BinaryOperationMultiplication:
		leftNum := toFloat(left)
		rightNum := toFloat(right)
		return leftNum * rightNum, nil
	case ast.BinaryOperationDivision:
		leftNum := toFloat(left)
		rightNum := toFloat(right)
		if rightNum == 0 {
			return nil, fmt.Errorf("division by zero")
		}
		return leftNum / rightNum, nil
	case ast.BinaryOperationIntDivision:
		leftNum := int(toFloat(left))
		rightNum := int(toFloat(right))
		if rightNum == 0 {
			return nil, fmt.Errorf("division by zero")
		}
		return leftNum / rightNum, nil
	case ast.BinaryOperationMod:
		leftNum := int(toFloat(left))
		rightNum := int(toFloat(right))
		if rightNum == 0 {
			return nil, fmt.Errorf("division by zero")
		}
		return leftNum % rightNum, nil
	case ast.BinaryOperationExponentiation:
		return math.Pow(toFloat(left), toFloat(right)), nil
	case ast.BinaryOperationEqual:
		return compareEqual(left, right), nil
	case ast.BinaryOperationNotEqual:
		return !compareEqual(left, right), nil
	case ast.BinaryOperationLess:
		return compareLess(left, right), nil
	case ast.BinaryOperationGreater:
		return compareLess(right, left), nil
	case ast.BinaryOperationLessOrEqual:
		return !compareLess(right, left), nil
	case ast.BinaryOperationGreaterOrEqual:
		return !compareLess(left, right), nil
	case ast.BinaryOperationConcatenation:
		return toString(left) + toString(right), nil
	case ast.BinaryOperationIs:
		return left == right, nil
	case ast.BinaryOperationXor, ast.BinaryOperationEqv, ast.BinaryOperationImp:
		// TODO: implement bitwise operations
		return nil, fmt.Errorf("binary operation %d not yet implemented", op)
	default:
		return nil, fmt.Errorf("unknown binary operator: %d", op)
	}
}

// compareEqual compares two values for equality
func compareEqual(left, right interface{}) bool {
	if left == nil && right == nil {
		return true
	}
	if left == nil || right == nil {
		return false
	}

	// Compare as strings first
	leftStr := toString(left)
	rightStr := toString(right)
	if leftStr == rightStr {
		return true
	}

	// Try numeric comparison
	if ln, lok := toNumeric(left); lok {
		if rn, rok := toNumeric(right); rok {
			return ln == rn
		}
	}

	return false
}

// compareLess compares if left is less than right
func compareLess(left, right interface{}) bool {
	leftNum, lok := toNumeric(left)
	rightNum, rok := toNumeric(right)

	if lok && rok {
		return leftNum < rightNum
	}

	// String comparison
	return toString(left) < toString(right)
}

// toNumeric attempts to convert a value to numeric type
func toNumeric(val interface{}) (float64, bool) {
	if val == nil {
		return 0, true
	}
	switch v := val.(type) {
	case int:
		return float64(v), true
	case float64:
		return v, true
	case string:
		if f, err := strconv.ParseFloat(v, 64); err == nil {
			return f, true
		}
		return 0, false
	case bool:
		if v {
			return -1, true
		}
		return 0, true
	default:
		return 0, false
	}
}

// populateRequestData fills a RequestObject with data from HTTP request
func populateRequestData(req *asp.RequestObject, r *http.Request) {
	// Parse form data
	r.ParseForm()

	// Set query string parameters
	for key, values := range r.URL.Query() {
		if len(values) > 0 {
			req.CallMethod("QueryString", key, values[0])
		}
	}

	// Set form parameters
	for key, values := range r.PostForm {
		if len(values) > 0 {
			req.CallMethod("Form", key, values[0])
		}
	}

	// Set cookies
	for _, cookie := range r.Cookies() {
		req.CallMethod("Cookies", cookie.Name, cookie.Value)
	}

	// Set server variables
	req.CallMethod("ServerVariables", "REQUEST_METHOD", r.Method)
	req.CallMethod("ServerVariables", "REQUEST_PATH", r.URL.Path)
	req.CallMethod("ServerVariables", "QUERY_STRING", r.URL.RawQuery)
	req.CallMethod("ServerVariables", "REMOTE_ADDR", r.RemoteAddr)
}
