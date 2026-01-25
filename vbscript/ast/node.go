package ast

// Node is the base interface for all AST nodes
type Node interface {
	GetRange() Range
	SetRange(Range)
	GetLocation() Location
	SetLocation(Location)
}

// BaseNode provides common functionality for all AST nodes
type BaseNode struct {
	Range    Range
	Location Location
}

// GetRange returns the range of the node
func (n *BaseNode) GetRange() Range {
	return n.Range
}

// SetRange sets the range of the node
func (n *BaseNode) SetRange(r Range) {
	n.Range = r
}

// GetLocation returns the location of the node
func (n *BaseNode) GetLocation() Location {
	return n.Location
}

// SetLocation sets the location of the node
func (n *BaseNode) SetLocation(l Location) {
	n.Location = l
}

// Statement is the base interface for all statements
type Statement interface {
	Node
	isStatement()
}

// BaseStatement provides common functionality for all statements
type BaseStatement struct {
	BaseNode
}

func (s *BaseStatement) isStatement() {}

// Expression is the base interface for all expressions
type Expression interface {
	Node
	isExpression()
}

// BaseExpression provides common functionality for all expressions
type BaseExpression struct {
	BaseNode
}

func (e *BaseExpression) isExpression() {}
