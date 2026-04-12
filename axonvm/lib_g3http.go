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
	"io"
	"net/http"
	"strings"
	"time"
)

type G3HTTP struct {
	vm *VM
}

// newG3HTTPObject instantiates the G3HTTP custom functions library.
func (vm *VM) newG3HTTPObject() Value {
	obj := &G3HTTP{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3httpItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet acts as a getter.
func (h *G3HTTP) DispatchPropertyGet(propertyName string) Value {
	return h.DispatchMethod(propertyName, nil)
}

// DispatchMethod provides O(1) string matching resolution for all custom HTTP functions.
func (h *G3HTTP) DispatchMethod(methodName string, args []Value) Value {
	funcLower := strings.ToLower(methodName)

	switch funcLower {
	case "fetch", "request":
		if len(args) < 1 {
			return NewEmpty()
		}
		url := args[0].String()
		httpMethod := "GET"
		bodyStr := ""

		if len(args) > 1 {
			httpMethod = strings.ToUpper(args[1].String())
		}
		if len(args) > 2 {
			bodyStr = args[2].String()
		}

		return h.executeRequest(url, httpMethod, bodyStr)
	}

	return NewEmpty()
}

func (h *G3HTTP) executeRequest(reqUrl, method, bodyStr string) Value {
	var bodyReader io.Reader
	if bodyStr != "" {
		bodyReader = strings.NewReader(bodyStr)
	}

	req, err := http.NewRequest(method, reqUrl, bodyReader)
	if err != nil {
		return NewEmpty()
	}

	if bodyStr != "" {
		req.Header.Set("Content-Type", "application/json")
	}

	client := &http.Client{Timeout: 10 * time.Second}
	resp, err := client.Do(req)
	if err != nil {
		return NewEmpty()
	}
	defer resp.Body.Close()

	data, err := io.ReadAll(resp.Body)
	if err != nil {
		return NewEmpty()
	}

	respString := string(data)

	contentType := resp.Header.Get("Content-Type")
	if strings.Contains(strings.ToLower(contentType), "application/json") {
		// Use the JSON library to parse it directly into VM structures
		jsonLib := &G3JSON{vm: h.vm}
		parsedVal := jsonLib.DispatchMethod("Parse", []Value{NewString(respString)})
		if parsedVal.Type != VTEmpty {
			return parsedVal
		}
	}

	return NewString(respString)
}
