package parser

import (
	"testing"
)

func TestTryCatchSemicolonTolerance(t *testing.T) {
	tt(t, func() {
		src := `try { throw "boom"; }; catch(e) {};`
		_, err := ParseFile(nil, "", src, 0)
		is(err, nil)
	})
}

func TestIfElseSemicolonTolerance(t *testing.T) {
	tt(t, func() {
		src := `if (true) {}; else if (false) {}; else {};`
		_, err := ParseFile(nil, "", src, 0)
		is(err, nil)
	})
}
