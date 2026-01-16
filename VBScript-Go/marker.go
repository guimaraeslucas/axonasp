package vbscript

// Marker represents a position in the source code
type Marker struct {
	Index  int // Absolute character position in source
	Line   int // Line number (1-based)
	Column int // Column number (0-based)
}

// NewMarker creates a new Marker with the given position information
func NewMarker(index, line, column int) Marker {
	return Marker{
		Index:  index,
		Line:   line,
		Column: column,
	}
}
