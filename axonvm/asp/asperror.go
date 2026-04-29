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
 *
 * Attribution Notice:
 * If this software is used in other projects, the name "AxonASP Server"
 * must be cited in the documentation or "About" section.
 *
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 */
package asp

import (
	"fmt"
	"strings"

	"g3pix.com.br/axonasp/jscript"
	"g3pix.com.br/axonasp/vbscript"
)

// ASPError captures details of runtime errors.
type ASPError struct {
	ASPCode        int
	ASPDescription string
	Number         int
	Source         string
	Description    string
	HelpFile       string
	HelpContext    int
	File           string
	Line           int
	Column         int
	Category       string
}

// NewASPError creates a new empty ASP error with default category.
func NewASPError() *ASPError {
	return (&ASPError{Category: "ASP"}).Normalize()
}

// Normalize fills empty ASPError fields with ASP-compatible defaults.
func (e *ASPError) Normalize() *ASPError {
	if e == nil {
		return &ASPError{Category: "ASP"}
	}

	if strings.TrimSpace(e.Category) == "" {
		e.Category = "ASP"
	}
	if strings.TrimSpace(e.Description) == "" {
		e.Description = strings.TrimSpace(e.ASPDescription)
	}
	if strings.TrimSpace(e.ASPDescription) == "" {
		e.ASPDescription = strings.TrimSpace(e.Description)
	}
	if strings.TrimSpace(e.Source) == "" {
		e.Source = e.Category
	}

	return e
}

// Clone returns a detached copy of the ASP error.
func (e *ASPError) Clone() *ASPError {
	if e == nil {
		return NewASPError()
	}

	clone := *e
	return (&clone).Normalize()
}

// Error formats the ASP error using a VBScript-compatible detail layout.
func (e *ASPError) Error() string {
	if e == nil {
		return ""
	}

	normalized := e.Clone()
	var builder strings.Builder
	builder.Grow(256)
	builder.WriteString(normalized.Source)
	if normalized.Number != 0 {
		builder.WriteString(" '")
		builder.WriteString(fmt.Sprintf("%08X", uint32(int32(normalized.Number))))
		builder.WriteString("'")
	}

	if normalized.Description != "" {
		builder.WriteString("\n")
		builder.WriteString(normalized.Description)
	}

	builder.WriteString("\nCategory: ")
	builder.WriteString(normalized.Category)
	builder.WriteString("\nColumn: ")
	builder.WriteString(fmt.Sprintf("%d", normalized.Column))
	builder.WriteString("\nDescription: ")
	builder.WriteString(normalized.Description)
	builder.WriteString("\nFile: ")
	builder.WriteString(normalized.File)
	builder.WriteString("\nLine: ")
	builder.WriteString(fmt.Sprintf("%d", normalized.Line))
	builder.WriteString("\nNumber: ")
	builder.WriteString(fmt.Sprintf("%d", normalized.Number))
	builder.WriteString("\nSource: ")
	builder.WriteString(normalized.Source)

	if normalized.ASPCode != 0 {
		builder.WriteString("\nASPCode: ")
		builder.WriteString(fmt.Sprintf("%d", normalized.ASPCode))
	}
	if normalized.ASPDescription != "" {
		builder.WriteString("\nASPDescription: ")
		builder.WriteString(normalized.ASPDescription)
	}

	return builder.String()
}

// NewVBScriptASPError creates an ASPError from a VBScript catalog code and execution context.
func NewVBScriptASPError(code vbscript.VBSyntaxErrorCode, source string, category string, description string, file string, line int, column int) *ASPError {
	detail := strings.TrimSpace(description)
	if detail == "" {
		detail = code.String()
	}

	err := &ASPError{
		ASPCode:        int(code),
		ASPDescription: detail,
		Number:         vbscript.HRESULTFromVBScriptCode(code),
		Source:         strings.TrimSpace(source),
		Description:    code.String(),
		File:           strings.TrimSpace(file),
		Line:           line,
		Column:         column,
		Category:       strings.TrimSpace(category),
	}

	if err.Source == "" {
		err.Source = "VBScript runtime error"
	}
	if err.Category == "" {
		err.Category = "VBScript runtime"
	}

	return err.Normalize()
}

// NewASPErrorFromVBSyntaxError converts a compiler syntax error into the ASPError object model.
func NewASPErrorFromVBSyntaxError(err *vbscript.VBSyntaxError) *ASPError {
	if err == nil {
		return NewASPError()
	}

	return (&ASPError{
		ASPCode:        err.ASPCode,
		ASPDescription: err.ASPDescription,
		Number:         err.Number,
		Source:         err.Source,
		Description:    err.Description,
		File:           err.File,
		Line:           err.Line,
		Column:         err.Column,
		Category:       err.Category,
	}).Normalize()
}

// NewASPErrorFromJSSyntaxError converts a JScript syntax error into the ASPError object model.
func NewASPErrorFromJSSyntaxError(err *jscript.JSSyntaxError) *ASPError {
	if err == nil {
		return NewASPError()
	}

	return (&ASPError{
		ASPCode:        err.ASPCode,
		ASPDescription: err.ASPDescription,
		Number:         err.Number,
		Source:         err.Source,
		Description:    err.Description,
		File:           err.File,
		Line:           err.Line,
		Column:         err.Column,
		Category:       err.Category,
	}).Normalize()
}

// NewASPErrorFromMessage creates a generic ASPError when no VBScript catalog mapping is available.
func NewASPErrorFromMessage(category string, source string, description string, file string, line int, column int) *ASPError {
	err := &ASPError{
		ASPDescription: strings.TrimSpace(description),
		Source:         strings.TrimSpace(source),
		Description:    strings.TrimSpace(description),
		File:           strings.TrimSpace(file),
		Line:           line,
		Column:         column,
		Category:       strings.TrimSpace(category),
	}

	return err.Normalize()
}
