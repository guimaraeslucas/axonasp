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
	"encoding/json"
	"fmt"
	"log"
	"os"
	"strings"
)

// G3JSON implements Component interface for JSON operations
type G3JSON struct{}

func (j *G3JSON) GetProperty(name string) any {
	return nil
}

func (j *G3JSON) SetProperty(name string, value any) {}

func (j *G3JSON) CallMethod(name string, args ...interface{}) interface{} {
	switch strings.ToLower(name) {
	case "parse":
		if len(args) > 0 {
			return j.Parse(fmt.Sprintf("%v", args[0]))
		}
	case "stringify":
		if len(args) > 0 {
			return j.Stringify(args[0])
		}
	case "newobject":
		return j.NewObject()
	case "newarray":
		return j.NewArray()
	case "loadfile":
		if len(args) > 0 {
			return j.LoadFile(fmt.Sprintf("%v", args[0]))
		}
	}
	return nil
}

func (j *G3JSON) Parse(jsonStr string) any {
	var result interface{}
	// Tenta fazer o unmarshal. Se for objeto vira map[string]interface{}, se array vira []interface{}
	err := json.Unmarshal([]byte(jsonStr), &result)
	if err != nil {
		return nil // Ou retornar um objeto de erro se preferir
	}
	return result
}

func (j *G3JSON) Stringify(data interface{}) string {
	bytes, err := json.Marshal(data)
	if err != nil {
		log.Printf("AxonASP JSON Error: %v\n", err) // Log no console do servidor
		return ""
	}
	return string(bytes)
}

func (j *G3JSON) NewObject() map[string]interface{} {
	return make(map[string]interface{})
}

func (j *G3JSON) NewArray() []interface{} {
	return make([]interface{}, 0)
}

// Bônus: Carregar direto de arquivo (muito comum em configs)
func (j *G3JSON) LoadFile(path string) interface{} {
	content, err := os.ReadFile(path)
	if err != nil {
		return nil
	}
	return j.Parse(string(content))
}
