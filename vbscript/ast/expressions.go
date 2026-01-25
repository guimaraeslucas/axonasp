package ast

import "time"

// Identifier represents a variable or function name
type Identifier struct {
	BaseExpression
	Name        string
	IsBracketed bool // True if the identifier was written with brackets [name] to escape reserved words
}

const IdentifierMaxLength = 255

// NewIdentifier creates a new Identifier expression
func NewIdentifier(name string) *Identifier {
	if name == "" {
		panic("identifier name cannot be empty")
	}
	return &Identifier{
		Name:        name,
		IsBracketed: false,
	}
}

// NewBracketedIdentifier creates a new bracketed Identifier expression
// Bracketed identifiers [name] are used to escape reserved words in VBScript
func NewBracketedIdentifier(name string) *Identifier {
	if name == "" {
		panic("identifier name cannot be empty")
	}
	return &Identifier{
		Name:        name,
		IsBracketed: true,
	}
}

// LiteralExpression is the base for all literal expressions
type LiteralExpression interface {
	Expression
	isLiteralExpression()
}

// BaseLiteralExpression provides common functionality for literals
type BaseLiteralExpression struct {
	BaseExpression
}

func (l *BaseLiteralExpression) isLiteralExpression() {}

// StringLiteral represents a string constant
type StringLiteral struct {
	BaseLiteralExpression
	Value string
}

// NewStringLiteral creates a new StringLiteral
func NewStringLiteral(value string) *StringLiteral {
	return &StringLiteral{
		Value: value,
	}
}

// IntegerLiteral represents an integer constant
type IntegerLiteral struct {
	BaseLiteralExpression
	Value int64
}

// NewIntegerLiteral creates a new IntegerLiteral
func NewIntegerLiteral(value int64) *IntegerLiteral {
	return &IntegerLiteral{
		Value: value,
	}
}

// FloatLiteral represents a floating-point constant
type FloatLiteral struct {
	BaseLiteralExpression
	Value float64
}

// NewFloatLiteral creates a new FloatLiteral
func NewFloatLiteral(value float64) *FloatLiteral {
	return &FloatLiteral{
		Value: value,
	}
}

// DateLiteral represents a date constant
type DateLiteral struct {
	BaseLiteralExpression
	Value time.Time
}

// NewDateLiteral creates a new DateLiteral
func NewDateLiteral(value time.Time) *DateLiteral {
	return &DateLiteral{
		Value: value,
	}
}

// BooleanLiteral represents true or false
type BooleanLiteral struct {
	BaseLiteralExpression
	Value bool
}

// NewBooleanLiteral creates a new BooleanLiteral
func NewBooleanLiteral(value bool) *BooleanLiteral {
	return &BooleanLiteral{
		Value: value,
	}
}

// NullLiteral represents Null
type NullLiteral struct {
	BaseLiteralExpression
}

// NewNullLiteral creates a new NullLiteral
func NewNullLiteral() *NullLiteral {
	return &NullLiteral{}
}

// EmptyLiteral represents Empty
type EmptyLiteral struct {
	BaseLiteralExpression
}

// NewEmptyLiteral creates a new EmptyLiteral
func NewEmptyLiteral() *EmptyLiteral {
	return &EmptyLiteral{}
}

// NothingLiteral represents Nothing
type NothingLiteral struct {
	BaseLiteralExpression
}

// NewNothingLiteral creates a new NothingLiteral
func NewNothingLiteral() *NothingLiteral {
	return &NothingLiteral{}
}

// UnaryOperation represents unary operations
type UnaryOperation int

const (
	UnaryOperationPlus UnaryOperation = iota
	UnaryOperationMinus
	UnaryOperationNot
)

// UnaryExpression represents unary operations like -x, +x, Not x
type UnaryExpression struct {
	BaseExpression
	Operation UnaryOperation
	Argument  Expression
}

// NewUnaryExpression creates a new UnaryExpression
func NewUnaryExpression(op UnaryOperation, arg Expression) *UnaryExpression {
	if arg == nil {
		panic("argument cannot be nil")
	}
	return &UnaryExpression{
		Operation: op,
		Argument:  arg,
	}
}

// BinaryOperation represents binary operations
type BinaryOperation int

const (
	BinaryOperationExponentiation BinaryOperation = iota
	BinaryOperationMultiplication
	BinaryOperationDivision
	BinaryOperationIntDivision
	BinaryOperationAddition
	BinaryOperationSubtraction
	BinaryOperationConcatenation
	BinaryOperationMod
	BinaryOperationIs
	BinaryOperationAnd
	BinaryOperationOr
	BinaryOperationXor
	BinaryOperationEqv
	BinaryOperationImp
	BinaryOperationEqual
	BinaryOperationNotEqual
	BinaryOperationLess
	BinaryOperationGreater
	BinaryOperationLessOrEqual
	BinaryOperationGreaterOrEqual
)

// BinaryExpression represents binary operations like a + b, a And b, etc.
type BinaryExpression struct {
	BaseExpression
	Operation BinaryOperation
	Left      Expression
	Right     Expression
}

// NewBinaryExpression creates a new BinaryExpression
func NewBinaryExpression(op BinaryOperation, left, right Expression) *BinaryExpression {
	if left == nil {
		panic("left operand cannot be nil")
	}
	if right == nil {
		panic("right operand cannot be nil")
	}
	return &BinaryExpression{
		Operation: op,
		Left:      left,
		Right:     right,
	}
}

// MemberExpression represents property access like obj.property
type MemberExpression struct {
	BaseExpression
	Object   Expression
	Property *Identifier
}

// NewMemberExpression creates a new MemberExpression
func NewMemberExpression(obj Expression, prop *Identifier) *MemberExpression {
	if obj == nil {
		panic("object cannot be nil")
	}
	if prop == nil {
		panic("property cannot be nil")
	}
	return &MemberExpression{
		Object:   obj,
		Property: prop,
	}
}

// IndexOrCallExpression represents array indexing or function calls like array(0) or func()
type IndexOrCallExpression struct {
	BaseExpression
	Object  Expression
	Indexes []Expression
}

// NewIndexOrCallExpression creates a new IndexOrCallExpression
func NewIndexOrCallExpression(obj Expression) *IndexOrCallExpression {
	if obj == nil {
		panic("object cannot be nil")
	}
	return &IndexOrCallExpression{
		Object:  obj,
		Indexes: []Expression{},
	}
}

// NewExpression represents the New operator like New ClassName
type NewExpression struct {
	BaseExpression
	Argument Expression
}

// NewNewExpression creates a new NewExpression
func NewNewExpression(arg Expression) *NewExpression {
	if arg == nil {
		panic("argument cannot be nil")
	}
	return &NewExpression{
		Argument: arg,
	}
}

// MissingValueExpression represents a missing parameter value in a function call
type MissingValueExpression struct {
	BaseExpression
}

// NewMissingValueExpression creates a new MissingValueExpression
func NewMissingValueExpression() *MissingValueExpression {
	return &MissingValueExpression{}
}

// WithMemberAccessExpression represents property access within a With block like .property
type WithMemberAccessExpression struct {
	BaseExpression
	Property *Identifier
}

// NewWithMemberAccessExpression creates a new WithMemberAccessExpression
func NewWithMemberAccessExpression(prop *Identifier) *WithMemberAccessExpression {
	if prop == nil {
		panic("property cannot be nil")
	}
	return &WithMemberAccessExpression{
		Property: prop,
	}
}
