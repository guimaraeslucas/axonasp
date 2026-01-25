/*
 * AxonASP Server - Version 1.0
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
package server

import (
	"bytes"
	"fmt"
	"html/template"
	"path/filepath"
	"strings"
)

// G3TEMPLATE implements Component interface for Template operations
type G3TEMPLATE struct {
	ctx *ExecutionContext
}

func (t *G3TEMPLATE) GetProperty(name string) interface{} {
	return nil
}

func (t *G3TEMPLATE) SetProperty(name string, value interface{}) {}

func (t *G3TEMPLATE) CallMethod(name string, args ...interface{}) interface{} {
	if len(args) < 1 {
		return "Error: Template path required"
	}

	getStr := func(i int) string {
		if i >= len(args) {
			return ""
		}
		return fmt.Sprintf("%v", args[i])
	}

	method := strings.ToLower(name)

	switch method {
	case "render":
		relPath := getStr(0)
		fullPath := t.ctx.Server_MapPath(relPath)

		rootDir, _ := filepath.Abs(t.ctx.RootDir)
		absPath, _ := filepath.Abs(fullPath)

		if !strings.HasPrefix(strings.ToLower(absPath), strings.ToLower(rootDir)) {
			return "Error: Access Denied"
		}

		var data interface{}
		if len(args) > 1 {
			data = args[1]
		}

		tmpl, err := template.ParseFiles(fullPath)
		if err != nil {
			return fmt.Sprintf("Error parsing template: %v", err)
		}

		var buf bytes.Buffer
		if err := tmpl.Execute(&buf, data); err != nil {
			return fmt.Sprintf("Error executing template: %v", err)
		}

		return buf.String()
	}
	return nil
}

// TemplateHelper handles Go template rendering (Legacy)
func TemplateHelper(method string, args []string, ctx *ExecutionContext) interface{} {
	lib := &G3TEMPLATE{ctx: ctx}

	var ifaceArgs []interface{}
	for _, a := range args {
		ifaceArgs = append(ifaceArgs, EvaluateExpression(a, ctx))
	}

	return lib.CallMethod(method, ifaceArgs)
}
