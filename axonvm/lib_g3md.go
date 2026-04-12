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
	"bytes"
	"strings"

	"github.com/yuin/goldmark"
	"github.com/yuin/goldmark/extension"
	"github.com/yuin/goldmark/parser"
	"github.com/yuin/goldmark/renderer"
	"github.com/yuin/goldmark/renderer/html"
)

// G3MD stores runtime options for Markdown to HTML conversion.
type G3MD struct {
	hardWraps bool
	unsafe    bool
}

// NewG3MD creates a new G3MD object with default options.
func NewG3MD() *G3MD {
	return &G3MD{hardWraps: false, unsafe: false}
}

// DispatchMethod executes methods and property Let/Set behavior for G3MD.
func (md *G3MD) DispatchMethod(methodName string, args []Value) Value {
	switch {
	case strings.EqualFold(methodName, "Process"):
		if len(args) < 1 {
			return NewString("")
		}
		return NewString(md.process(args[0].String()))
	case strings.EqualFold(methodName, "HardWraps"):
		if len(args) > 0 {
			md.hardWraps = g3mdValueToBool(args[0])
			return Value{Type: VTEmpty}
		}
		return NewBool(md.hardWraps)
	case strings.EqualFold(methodName, "Unsafe"):
		if len(args) > 0 {
			md.unsafe = g3mdValueToBool(args[0])
			return Value{Type: VTEmpty}
		}
		return NewBool(md.unsafe)
	default:
		return Value{Type: VTEmpty}
	}
}

// DispatchPropertyGet resolves property reads for G3MD.
func (md *G3MD) DispatchPropertyGet(propertyName string) Value {
	switch {
	case strings.EqualFold(propertyName, "HardWraps"):
		return NewBool(md.hardWraps)
	case strings.EqualFold(propertyName, "Unsafe"):
		return NewBool(md.unsafe)
	default:
		return Value{Type: VTEmpty}
	}
}

// DispatchPropertySet handles Let assignments on G3MD properties.
func (md *G3MD) DispatchPropertySet(propertyName string, val Value) {
	switch {
	case strings.EqualFold(propertyName, "HardWraps"):
		md.hardWraps = g3mdValueToBool(val)
	case strings.EqualFold(propertyName, "Unsafe"):
		md.unsafe = g3mdValueToBool(val)
	}
}

// process converts Markdown source into HTML using GFM extensions.
func (md *G3MD) process(source string) string {
	options := make([]renderer.Option, 0, 2)
	if md.hardWraps {
		options = append(options, html.WithHardWraps())
	}
	if md.unsafe {
		options = append(options, html.WithUnsafe())
	}

	markdown := goldmark.New(
		goldmark.WithExtensions(extension.GFM),
		goldmark.WithParserOptions(parser.WithAutoHeadingID()),
		goldmark.WithRendererOptions(options...),
	)

	var buffer bytes.Buffer
	if err := markdown.Convert([]byte(source), &buffer); err != nil {
		return ""
	}

	return buffer.String()
}

// g3mdValueToBool coerces VM values to bool for G3MD property assignment.
func g3mdValueToBool(value Value) bool {
	switch value.Type {
	case VTBool, VTInteger, VTNativeObject, VTBuiltin:
		return value.Num != 0
	case VTDouble:
		return value.Flt != 0
	case VTString:
		text := strings.TrimSpace(value.Str)
		if strings.EqualFold(text, "false") || text == "" || text == "0" {
			return false
		}
		return true
	default:
		return false
	}
}
