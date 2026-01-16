package ast

import "fmt"

// Position represents a line and column in source code
type Position struct {
	Line   int
	Column int
}

// NewPosition creates a new Position
func NewPosition(line, column int) Position {
	return Position{
		Line:   line,
		Column: column,
	}
}

// String returns the string representation of Position
func (p Position) String() string {
	return fmt.Sprintf("%d,%d", p.Line, p.Column)
}

// Range represents a range in source code from Start to End
type Range struct {
	Start int
	End   int
}

// NewRange creates a new Range
func NewRange(start, end int) Range {
	return Range{
		Start: start,
		End:   end,
	}
}

// String returns the string representation of Range
func (r Range) String() string {
	return fmt.Sprintf("[%d..%d)", r.Start, r.End)
}

// Location represents a location in source code with start and end positions
type Location struct {
	Start Position
	End   Position
}

// NewLocation creates a new Location
func NewLocation(start, end Position) Location {
	return Location{
		Start: start,
		End:   end,
	}
}

// String returns the string representation of Location
func (l Location) String() string {
	return fmt.Sprintf("%s...%s", l.Start, l.End)
}
