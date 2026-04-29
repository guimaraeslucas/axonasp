/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 */
package jscript

import (
	"strconv"
	"strings"
)

// JSSyntaxError represents a JScript compilation error.
type JSSyntaxError struct {
	Code           JSSyntaxErrorCode
	Line           int
	Column         int
	ASPCode        int
	ASPDescription string
	Category       string
	Description    string
	File           string
	Number         int
	Source         string
}

// NewJSSyntaxError creates a new JScript syntax error instance.
func NewJSSyntaxError(code JSSyntaxErrorCode, line, column int) *JSSyntaxError {
	description := code.String()
	return &JSSyntaxError{
		Code:           code,
		Line:           line,
		Column:         normalizeJScriptColumn(column),
		ASPCode:        int(code),
		ASPDescription: description,
		Category:       "JScript compilation",
		Description:    description,
		File:           "",
		Number:         HRESULTFromJScriptCode(code),
		Source:         "JScript compilation error",
	}
}

// WithFile attaches a source file path to the syntax error.
func (e *JSSyntaxError) WithFile(file string) *JSSyntaxError {
	if e == nil {
		return nil
	}

	e.File = strings.TrimSpace(file)
	return e
}

// WithASPDescription overrides the ASP description while keeping catalog text intact.
func (e *JSSyntaxError) WithASPDescription(description string) *JSSyntaxError {
	if e == nil {
		return nil
	}

	trimmed := strings.TrimSpace(description)
	if trimmed != "" {
		e.ASPDescription = trimmed
	}
	return e
}

// Error implements the error interface.
func (e *JSSyntaxError) Error() string {
	if e == nil {
		return ""
	}

	var builder strings.Builder
	builder.Grow(224)
	builder.WriteString(e.Source)
	if e.Code != 0 {
		builder.WriteString(" '")
		builder.WriteString(HRESULTHexFromJScriptCode(e.Code))
		builder.WriteString("'")
	}

	if strings.TrimSpace(e.Description) != "" {
		builder.WriteString("\n")
		builder.WriteString(e.Description)
	}

	builder.WriteString("\nCategory: ")
	builder.WriteString(e.Category)
	builder.WriteString("\nColumn: ")
	builder.WriteString(strconv.Itoa(e.Column))
	builder.WriteString("\nDescription: ")
	builder.WriteString(e.Description)
	builder.WriteString("\nFile: ")
	builder.WriteString(e.File)
	builder.WriteString("\nLine: ")
	builder.WriteString(strconv.Itoa(e.Line))
	builder.WriteString("\nNumber: ")
	builder.WriteString(strconv.Itoa(e.Number))
	builder.WriteString("\nSource: ")
	builder.WriteString(e.Source)

	return builder.String()
}

// normalizeJScriptColumn converts parser column to a non-negative value.
func normalizeJScriptColumn(column int) int {
	if column < 0 {
		return 0
	}

	return column
}
