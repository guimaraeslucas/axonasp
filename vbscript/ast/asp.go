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
package ast

// HTMLStatement represents a block of HTML code in an ASP file
type HTMLStatement struct {
	BaseStatement
	Content string
}

// NewHTMLStatement creates a new HTMLStatement
func NewHTMLStatement(content string) *HTMLStatement {
	return &HTMLStatement{
		Content: content,
	}
}

// ASPExpressionStatement represents <%= expression %>
type ASPExpressionStatement struct {
	BaseStatement
	Expression Expression
}

// NewASPExpressionStatement creates a new ASPExpressionStatement
func NewASPExpressionStatement(expr Expression) *ASPExpressionStatement {
	if expr == nil {
		panic("expression cannot be nil")
	}
	return &ASPExpressionStatement{
		Expression: expr,
	}
}

// ASPDirectiveStatement represents <%@ directive %>
type ASPDirectiveStatement struct {
	BaseStatement
	Attributes map[string]string
}

// NewASPDirectiveStatement creates a new ASPDirectiveStatement
func NewASPDirectiveStatement() *ASPDirectiveStatement {
	return &ASPDirectiveStatement{
		Attributes: make(map[string]string),
	}
}

// IncludeStatement represents <!--#include file/virtual="..."-->
type IncludeStatement struct {
	BaseStatement
	Virtual bool
	Path    string
}

// NewIncludeStatement creates a new IncludeStatement
func NewIncludeStatement(virtual bool, path string) *IncludeStatement {
	return &IncludeStatement{
		Virtual: virtual,
		Path:    path,
	}
}
