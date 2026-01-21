package vbscript

import (
	"fmt"
	"testing"
)

func TestWherePanicHappens(t *testing.T) {
	defer func() {
		if r := recover(); r != nil {
			fmt.Printf("PANIC: %v\n", r)
		}
	}()

	code := "x=1 : response.end"
	parser := NewParser(code)
	program := parser.Parse()
	
	if program != nil {
		fmt.Printf("Success!\n")
	}
}
