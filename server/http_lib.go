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
	"fmt"
	"io"
	"net/http"
	"strings"
	"time"
)

// G3HTTP implements Component interface for HTTP operations
type G3HTTP struct {
	ctx *ExecutionContext
}

func (h *G3HTTP) GetProperty(name string) interface{} {
	return nil
}

func (h *G3HTTP) SetProperty(name string, value interface{}) {}

func (h *G3HTTP) CallMethod(name string, args ...interface{}) interface{} {
	if len(args) < 1 {
		return nil
	}

	method := strings.ToLower(name)

	switch method {
	case "fetch", "request":
		// Args: URL, [Method], [Body]
		url := fmt.Sprintf("%v", args[0])
		httpMethod := "GET"
		bodyStr := ""

		if len(args) > 1 {
			httpMethod = strings.ToUpper(fmt.Sprintf("%v", args[1]))
		}
		if len(args) > 2 {
			bodyStr = fmt.Sprintf("%v", args[2])
		}

		return h.executeRequest(url, httpMethod, bodyStr)
	}
	return nil
}

func (h *G3HTTP) executeRequest(url, method, bodyStr string) interface{} {
	var bodyReader io.Reader
	if bodyStr != "" {
		bodyReader = strings.NewReader(bodyStr)
	}

	req, err := http.NewRequest(method, url, bodyReader)
	if err != nil {
		return nil
	}

	if bodyStr != "" {
		req.Header.Set("Content-Type", "application/json")
	}

	client := &http.Client{Timeout: 10 * time.Second}
	resp, err := client.Do(req)
	if err != nil {
		return nil
	}
	defer resp.Body.Close()

	data, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil
	}

	respString := string(data)

	contentType := resp.Header.Get("Content-Type")
	if strings.Contains(strings.ToLower(contentType), "application/json") {
		lib := &G3JSON{}
		parsed := lib.Parse(respString)
		if parsed != nil {
			// Convert map[string]interface{} to DictionaryLibrary for VBScript compatibility
			return h.mapToDictionary(parsed)
		}
	}

	return respString
}

// mapToDictionary converts Go map or slice to VBScript-compatible Dictionary/Array
func (h *G3HTTP) mapToDictionary(data interface{}) interface{} {
	switch v := data.(type) {
	case map[string]interface{}:
		// Return map directly for VBScript subscript access
		result := make(map[string]interface{})
		for key, value := range v {
			result[key] = h.mapToDictionary(value)
		}
		return result
	case []interface{}:
		// Convert array recursively
		result := make([]interface{}, len(v))
		for i, item := range v {
			result[i] = h.mapToDictionary(item)
		}
		return result
	default:
		return data
	}
}

// FetchHelper for backward compatibility
func FetchHelper(args []string, ctx *ExecutionContext) interface{} {
	if len(args) < 1 {
		return nil
	}

	lib := &G3HTTP{ctx: ctx}

	// Evaluate args
	var ifaceArgs []interface{}
	for _, a := range args {
		ifaceArgs = append(ifaceArgs, EvaluateExpression(a, ctx))
	}

	return lib.CallMethod("fetch", ifaceArgs)
}
