package ast

// AccessModifiers

// MethodAccessModifier represents access level for methods
type MethodAccessModifier int

const (
	MethodAccessModifierNone MethodAccessModifier = iota
	MethodAccessModifierPublic
	MethodAccessModifierPrivate
	MethodAccessModifierPublicDefault
)

// MemberAccessModifier represents access level for members
type MemberAccessModifier int

const (
	MemberAccessModifierNone MemberAccessModifier = iota
	MemberAccessModifierPublic
	MemberAccessModifierPrivate
)

// FieldAccessModifier represents access level for fields
type FieldAccessModifier int

const (
	FieldAccessModifierNone FieldAccessModifier = iota
	FieldAccessModifierPublic
	FieldAccessModifierPrivate
)

// LoopType represents the type of loop
type LoopType int

const (
	LoopTypeNone LoopType = iota
	LoopTypeWhile
	LoopTypeUntil
)

// ConditionTestType represents when condition is tested
type ConditionTestType int

const (
	ConditionTestTypeNone ConditionTestType = iota
	ConditionTestTypePreTest
	ConditionTestTypePostTest
)

// ExitStatement is the base for all Exit statements
type ExitStatement interface {
	Statement
	isExitStatement()
}

// BaseExitStatement provides common functionality for exit statements
type BaseExitStatement struct {
	BaseStatement
}

func (e *BaseExitStatement) isExitStatement() {}

// ExitDoStatement represents Exit Do
type ExitDoStatement struct {
	BaseExitStatement
}

// NewExitDoStatement creates a new ExitDoStatement
func NewExitDoStatement() *ExitDoStatement {
	return &ExitDoStatement{}
}

// ExitForStatement represents Exit For
type ExitForStatement struct {
	BaseExitStatement
}

// NewExitForStatement creates a new ExitForStatement
func NewExitForStatement() *ExitForStatement {
	return &ExitForStatement{}
}

// ExitSubStatement represents Exit Sub
type ExitSubStatement struct {
	BaseExitStatement
}

// NewExitSubStatement creates a new ExitSubStatement
func NewExitSubStatement() *ExitSubStatement {
	return &ExitSubStatement{}
}

// ExitFunctionStatement represents Exit Function
type ExitFunctionStatement struct {
	BaseExitStatement
}

// NewExitFunctionStatement creates a new ExitFunctionStatement
func NewExitFunctionStatement() *ExitFunctionStatement {
	return &ExitFunctionStatement{}
}

// ExitPropertyStatement represents Exit Property
type ExitPropertyStatement struct {
	BaseExitStatement
}

// NewExitPropertyStatement creates a new ExitPropertyStatement
func NewExitPropertyStatement() *ExitPropertyStatement {
	return &ExitPropertyStatement{}
}

// AssignmentStatement represents an assignment like x = 5
type AssignmentStatement struct {
	BaseStatement
	Left  Expression
	Right Expression
	Set   bool // true if Set statement
}

// NewAssignmentStatement creates a new AssignmentStatement
func NewAssignmentStatement(left, right Expression, setStatement bool) *AssignmentStatement {
	if left == nil {
		panic("left operand cannot be nil")
	}
	if right == nil {
		panic("right operand cannot be nil")
	}
	return &AssignmentStatement{
		Left:  left,
		Right: right,
		Set:   setStatement,
	}
}

// CallStatement represents a Call statement
type CallStatement struct {
	BaseStatement
	Callee Expression
}

// NewCallStatement creates a new CallStatement
func NewCallStatement(callee Expression) *CallStatement {
	if callee == nil {
		panic("callee cannot be nil")
	}
	return &CallStatement{
		Callee: callee,
	}
}

// CallSubStatement represents a subroutine call
type CallSubStatement struct {
	BaseStatement
	Callee    Expression
	Arguments []Expression
}

// NewCallSubStatement creates a new CallSubStatement
func NewCallSubStatement(callee Expression) *CallSubStatement {
	if callee == nil {
		panic("callee cannot be nil")
	}
	return &CallSubStatement{
		Callee:    callee,
		Arguments: []Expression{},
	}
}

// EraseStatement represents an Erase statement
type EraseStatement struct {
	BaseStatement
	Identifier *Identifier
}

// NewEraseStatement creates a new EraseStatement
func NewEraseStatement(id *Identifier) *EraseStatement {
	if id == nil {
		panic("identifier cannot be nil")
	}
	return &EraseStatement{
		Identifier: id,
	}
}

// OnErrorResumeNextStatement represents On Error Resume Next
type OnErrorResumeNextStatement struct {
	BaseStatement
}

// NewOnErrorResumeNextStatement creates a new OnErrorResumeNextStatement
func NewOnErrorResumeNextStatement() *OnErrorResumeNextStatement {
	return &OnErrorResumeNextStatement{}
}

// OnErrorGoTo0Statement represents On Error GoTo 0
type OnErrorGoTo0Statement struct {
	BaseStatement
}

// NewOnErrorGoTo0Statement creates a new OnErrorGoTo0Statement
func NewOnErrorGoTo0Statement() *OnErrorGoTo0Statement {
	return &OnErrorGoTo0Statement{}
}

// IfStatement represents an If...Then...Else statement
type IfStatement struct {
	BaseStatement
	Test       Expression
	Consequent Statement
	Alternate  Statement
}

// NewIfStatement creates a new IfStatement
func NewIfStatement(test Expression, consequent Statement, alternate Statement) *IfStatement {
	if test == nil {
		panic("test condition cannot be nil")
	}
	if consequent == nil {
		panic("consequent cannot be nil")
	}
	return &IfStatement{
		Test:       test,
		Consequent: consequent,
		Alternate:  alternate,
	}
}

// ElseIfStatement represents an ElseIf statement
type ElseIfStatement struct {
	BaseStatement
	Test       Expression
	Consequent Statement
	Alternate  Statement
}

// NewElseIfStatement creates a new ElseIfStatement
func NewElseIfStatement(test Expression, consequent Statement, alternate Statement) *ElseIfStatement {
	if test == nil {
		panic("test condition cannot be nil")
	}
	if consequent == nil {
		panic("consequent cannot be nil")
	}
	return &ElseIfStatement{
		Test:       test,
		Consequent: consequent,
		Alternate:  alternate,
	}
}

// ForStatement represents a For...Next statement
type ForStatement struct {
	BaseStatement
	Identifier *Identifier
	From       Expression
	To         Expression
	Step       Expression
	Body       []Statement
}

// NewForStatement creates a new ForStatement
func NewForStatement(id *Identifier, from, to, step Expression) *ForStatement {
	if id == nil {
		panic("identifier cannot be nil")
	}
	if from == nil {
		panic("from expression cannot be nil")
	}
	if to == nil {
		panic("to expression cannot be nil")
	}
	return &ForStatement{
		Identifier: id,
		From:       from,
		To:         to,
		Step:       step,
		Body:       []Statement{},
	}
}

// ForEachStatement represents a For Each...Next statement
type ForEachStatement struct {
	BaseStatement
	Identifier *Identifier
	In         Expression
	Body       []Statement
}

// NewForEachStatement creates a new ForEachStatement
func NewForEachStatement(id *Identifier, in Expression) *ForEachStatement {
	if id == nil {
		panic("identifier cannot be nil")
	}
	if in == nil {
		panic("in expression cannot be nil")
	}
	return &ForEachStatement{
		Identifier: id,
		In:         in,
		Body:       []Statement{},
	}
}

// DoStatement represents a Do...Loop statement
type DoStatement struct {
	BaseStatement
	LoopType  LoopType
	TestType  ConditionTestType
	Condition Expression
	Body      []Statement
}

// NewDoStatement creates a new DoStatement
func NewDoStatement(loopType LoopType, testType ConditionTestType, condition Expression) *DoStatement {
	if condition == nil {
		panic("condition cannot be nil")
	}
	return &DoStatement{
		LoopType:  loopType,
		TestType:  testType,
		Condition: condition,
		Body:      []Statement{},
	}
}

// WhileStatement represents a While...Wend statement
type WhileStatement struct {
	BaseStatement
	Condition Expression
	Body      []Statement
}

// NewWhileStatement creates a new WhileStatement
func NewWhileStatement(condition Expression) *WhileStatement {
	if condition == nil {
		panic("condition cannot be nil")
	}
	return &WhileStatement{
		Condition: condition,
		Body:      []Statement{},
	}
}

// SelectStatement represents a Select Case statement
type SelectStatement struct {
	BaseStatement
	Condition Expression
	Cases     []*CaseStatement
}

// NewSelectStatement creates a new SelectStatement
func NewSelectStatement(condition Expression) *SelectStatement {
	if condition == nil {
		panic("condition cannot be nil")
	}
	return &SelectStatement{
		Condition: condition,
		Cases:     []*CaseStatement{},
	}
}

// CaseStatement represents a Case in a Select statement
type CaseStatement struct {
	BaseStatement
	Values []Expression
	Body   []Statement
}

// NewCaseStatement creates a new CaseStatement
func NewCaseStatement() *CaseStatement {
	return &CaseStatement{
		Values: []Expression{},
		Body:   []Statement{},
	}
}

// WithStatement represents a With...End With statement
type WithStatement struct {
	BaseStatement
	Expression Expression
	Body       []Statement
}

// NewWithStatement creates a new WithStatement
func NewWithStatement(expr Expression) *WithStatement {
	if expr == nil {
		panic("expression cannot be nil")
	}
	return &WithStatement{
		Expression: expr,
		Body:       []Statement{},
	}
}

// VariablesDeclaration represents Dim declarations
type VariablesDeclaration struct {
	BaseStatement
	Variables []*VariableDeclaration
}

// NewVariablesDeclaration creates a new VariablesDeclaration
func NewVariablesDeclaration() *VariablesDeclaration {
	return &VariablesDeclaration{
		Variables: []*VariableDeclaration{},
	}
}

// ConstsDeclaration represents Const declarations
type ConstsDeclaration struct {
	BaseStatement
	Modifier     MemberAccessModifier
	Declarations []*ConstDeclaration
}

// NewConstsDeclaration creates a new ConstsDeclaration
func NewConstsDeclaration(modifier MemberAccessModifier) *ConstsDeclaration {
	return &ConstsDeclaration{
		Modifier:     modifier,
		Declarations: []*ConstDeclaration{},
	}
}

// FieldsDeclaration represents field declarations in a class
type FieldsDeclaration struct {
	BaseStatement
	Modifier FieldAccessModifier
	Fields   []*FieldDeclaration
}

// NewFieldsDeclaration creates a new FieldsDeclaration
func NewFieldsDeclaration(modifier FieldAccessModifier) *FieldsDeclaration {
	return &FieldsDeclaration{
		Modifier: modifier,
		Fields:   []*FieldDeclaration{},
	}
}

// ReDimStatement represents a ReDim statement
type ReDimStatement struct {
	BaseStatement
	Preserve bool
	ReDims   []*ReDimDeclaration
}

// NewReDimStatement creates a new ReDimStatement
func NewReDimStatement(preserve bool) *ReDimStatement {
	return &ReDimStatement{
		Preserve: preserve,
		ReDims:   []*ReDimDeclaration{},
	}
}

// StatementList represents a list of statements
type StatementList struct {
	BaseStatement
	Statements []Statement
}

// NewStatementList creates a new StatementList
func NewStatementList() *StatementList {
	return &StatementList{
		Statements: []Statement{},
	}
}

// Add appends a statement to the list
func (sl *StatementList) Add(stmt Statement) {
	if stmt != nil {
		sl.Statements = append(sl.Statements, stmt)
	}
}

// Count returns the number of statements
func (sl *StatementList) Count() int {
	return len(sl.Statements)
}

// Get returns the statement at index
func (sl *StatementList) Get(index int) Statement {
	if index < 0 || index >= len(sl.Statements) {
		return nil
	}
	return sl.Statements[index]
}
