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
package axonvm

import (
	"fmt"
	"strings"

	"g3pix.com.br/axonasp/vbscript"
)

// compileASPDirective compiles one <%@ ... %> directive block into runtime directive opcodes.
func (c *Compiler) compileASPDirective() {
	for !c.matchEof() {
		c.skipDirectiveTrivia()

		if _, ok := c.next.(*vbscript.ASPCodeEndToken); ok {
			c.move()
			return
		}

		name := c.expectDirectiveIdentifier()
		c.expectDirectiveEquals()
		value := c.readDirectiveValue()
		c.validateASPDirective(name, value)

		nameIdx := c.addConstant(NewString(name))
		valueIdx := c.addConstant(NewString(value))
		c.emit(OpSetDirective, nameIdx, valueIdx)
	}

	panic(c.vbCompileError(vbscript.SyntaxError, "unterminated ASP directive block"))
}

// skipDirectiveTrivia consumes separators allowed inside ASP directive blocks.
func (c *Compiler) skipDirectiveTrivia() {
	for {
		switch c.next.(type) {
		case *vbscript.LineTerminationToken, *vbscript.ColonLineTerminationToken, *vbscript.CommentToken:
			c.move()
		default:
			return
		}
	}
}

// expectDirectiveIdentifier reads one directive attribute name.
func (c *Compiler) expectDirectiveIdentifier() string {
	switch t := c.next.(type) {
	case *vbscript.IdentifierToken:
		c.move()
		return t.Name
	case *vbscript.KeywordOrIdentifierToken:
		c.move()
		return t.Name
	case *vbscript.KeywordToken:
		c.move()
		return t.Keyword.String()
	default:
		panic(c.vbCompileError(vbscript.ExpectedIdentifier, fmt.Sprintf("expected directive identifier, got %T", c.next)))
	}
}

// expectDirectiveEquals consumes the equals sign used by ASP directives.
func (c *Compiler) expectDirectiveEquals() {
	if token, ok := c.next.(*vbscript.PunctuationToken); ok && token.Type == vbscript.PunctEqual {
		c.move()
		return
	}
	panic(c.vbCompileError(vbscript.ExpectedEqual, fmt.Sprintf("expected '=' in directive, got %T", c.next)))
}

// readDirectiveValue reads one ASP directive value and normalizes it into string form.
func (c *Compiler) readDirectiveValue() string {
	switch t := c.next.(type) {
	case *vbscript.StringLiteralToken:
		c.move()
		return t.Value
	case *vbscript.IdentifierToken:
		c.move()
		return t.Name
	case *vbscript.KeywordOrIdentifierToken:
		c.move()
		return t.Name
	case *vbscript.KeywordToken:
		c.move()
		return t.Keyword.String()
	case *vbscript.DecIntegerLiteralToken:
		c.move()
		return fmt.Sprintf("%d", t.Value)
	case *vbscript.OctIntegerLiteralToken:
		c.move()
		return fmt.Sprintf("%d", t.Value)
	case *vbscript.HexIntegerLiteralToken:
		c.move()
		return fmt.Sprintf("%d", t.Value)
	case *vbscript.TrueLiteralToken:
		c.move()
		return "True"
	case *vbscript.FalseLiteralToken:
		c.move()
		return "False"
	default:
		panic(c.vbCompileError(vbscript.ExpectedLiteral, fmt.Sprintf("expected directive value, got %T", c.next)))
	}
}

// validateASPDirective enforces compatibility rules for supported ASP directives.
func (c *Compiler) validateASPDirective(name string, value string) {
	switch strings.ToLower(name) {
	case "language":
		if !strings.EqualFold(strings.TrimSpace(value), "vbscript") {
			panic(c.vbCompileError(vbscript.SyntaxError, fmt.Sprintf("unsupported ASP language directive: %s", value)))
		}
	case "codepage":
		if !isDirectiveInteger(value) {
			panic(c.vbCompileError(vbscript.InvalidNumber, fmt.Sprintf("invalid ASP code page directive value: %s", value)))
		}
	case "enablesessionstate":
		normalized := strings.ToLower(strings.TrimSpace(value))
		if normalized != "true" && normalized != "false" && normalized != "readonly" {
			panic(c.vbCompileError(vbscript.SyntaxError, fmt.Sprintf("invalid ASP EnableSessionState directive value: %s", value)))
		}
	}
}

// isDirectiveInteger reports whether a directive value contains a base-10 integer.
func isDirectiveInteger(value string) bool {
	value = strings.TrimSpace(value)
	if value == "" {
		return false
	}
	for _, ch := range value {
		if ch < '0' || ch > '9' {
			return false
		}
	}
	return true
}
