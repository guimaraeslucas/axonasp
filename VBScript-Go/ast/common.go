package ast

// Program represents the root node of a VBScript AST
type Program struct {
	BaseNode
	OptionExplicit bool
	OptionCompare  OptionCompareMode
	OptionBase     int
	Body           []Statement
	Comments       []*Comment
}

// NewProgram creates a new Program node
func NewProgram(optionExplicit bool, optionCompare OptionCompareMode, optionBase int) *Program {
	return &Program{
		OptionExplicit: optionExplicit,
		OptionCompare:  optionCompare,
		OptionBase:     optionBase,
		Body:           []Statement{},
		Comments:       []*Comment{},
	}
}

// OptionCompareMode defines how string comparisons are evaluated
type OptionCompareMode int

const (
	OptionCompareBinary OptionCompareMode = iota
	OptionCompareText
)

func (p *Program) isStatement() {}

// CommentType represents the type of comment
type CommentType int

const (
	CommentTypeRem CommentType = iota
	CommentTypeSingleQuote
)

// Comment represents a comment in source code
type Comment struct {
	Type     CommentType
	Text     string
	Range    Range
	Location Location
}

// NewComment creates a new Comment
func NewComment(ctype CommentType, text string) *Comment {
	return &Comment{
		Type: ctype,
		Text: text,
	}
}

// String returns the comment type as string
func (ct CommentType) String() string {
	switch ct {
	case CommentTypeRem:
		return "REM"
	case CommentTypeSingleQuote:
		return "'"
	default:
		return "Unknown"
	}
}

// ParameterModifier represents how a parameter is passed
type ParameterModifier int

const (
	ParameterModifierNone ParameterModifier = iota
	ParameterModifierByRef
	ParameterModifierByVal
)

// Parameter represents a function/sub parameter
type Parameter struct {
	BaseNode
	Modifier    ParameterModifier
	Parentheses bool
	Identifier  *Identifier
}

// NewParameter creates a new Parameter
func NewParameter(id *Identifier, modifier ParameterModifier, parentheses bool) *Parameter {
	if id == nil {
		panic("identifier cannot be nil")
	}
	return &Parameter{
		Modifier:    modifier,
		Parentheses: parentheses,
		Identifier:  id,
	}
}

func (p *Parameter) isStatement() {}

// ConstDeclaration represents a constant declaration
type ConstDeclaration struct {
	BaseNode
	Identifier *Identifier
	Init       Expression
}

// NewConstDeclaration creates a new ConstDeclaration
func NewConstDeclaration(id *Identifier, init Expression) *ConstDeclaration {
	if id == nil {
		panic("identifier cannot be nil")
	}
	if init == nil {
		panic("initialization expression cannot be nil")
	}
	return &ConstDeclaration{
		Identifier: id,
		Init:       init,
	}
}

func (c *ConstDeclaration) isStatement() {}

// ReDimDeclaration represents a ReDim declaration
type ReDimDeclaration struct {
	BaseNode
	Identifier *Identifier
	ArrayDims  []Expression
}

// NewReDimDeclaration creates a new ReDimDeclaration
func NewReDimDeclaration(id *Identifier) *ReDimDeclaration {
	if id == nil {
		panic("identifier cannot be nil")
	}
	return &ReDimDeclaration{
		Identifier: id,
		ArrayDims:  []Expression{},
	}
}

func (r *ReDimDeclaration) isStatement() {}

// VariableDeclarationNode is the base interface for variable declarations
type VariableDeclarationNode interface {
	Node
	isVariableDeclarationNode()
}

// BaseVariableDeclarationNode provides common functionality for variable declarations
type BaseVariableDeclarationNode struct {
	BaseNode
	Identifier     *Identifier
	IsDynamicArray bool
	ArrayDims      []Expression
}

func (v *BaseVariableDeclarationNode) isVariableDeclarationNode() {}
func (v *BaseVariableDeclarationNode) isStatement()               {}

// VariableDeclaration represents a variable declaration
type VariableDeclaration struct {
	BaseVariableDeclarationNode
}

// NewVariableDeclaration creates a new VariableDeclaration
func NewVariableDeclaration(id *Identifier, isDynamicArray bool) *VariableDeclaration {
	if id == nil {
		panic("identifier cannot be nil")
	}
	return &VariableDeclaration{
		BaseVariableDeclarationNode: BaseVariableDeclarationNode{
			Identifier:     id,
			IsDynamicArray: isDynamicArray,
			ArrayDims:      []Expression{},
		},
	}
}

// FieldDeclaration represents a field declaration in a class
type FieldDeclaration struct {
	BaseVariableDeclarationNode
}

// NewFieldDeclaration creates a new FieldDeclaration
func NewFieldDeclaration(id *Identifier, isDynamicArray bool) *FieldDeclaration {
	if id == nil {
		panic("identifier cannot be nil")
	}
	return &FieldDeclaration{
		BaseVariableDeclarationNode: BaseVariableDeclarationNode{
			Identifier:     id,
			IsDynamicArray: isDynamicArray,
			ArrayDims:      []Expression{},
		},
	}
}
