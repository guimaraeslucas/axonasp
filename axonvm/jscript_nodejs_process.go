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
	"os"
	"strings"
)

// jsCreateProcessObject allocates the Node.js-compatible process global object.
// This object provides access to environment variables, command-line arguments,
// and process control methods like exit().
func (vm *VM) jsCreateProcessObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 10)

	// Type marker
	obj["__js_type"] = NewString("process")

	// Create process.env as a JS object mapping environment variables
	envObj := vm.jsCreateEnvObject()
	obj["env"] = envObj

	// Create process.argv as an array of command-line arguments
	argvArray := vm.jsCreateArgvArray()
	obj["argv"] = argvArray

	// Store function proxies for cwd and exit methods
	// These are special marker values that will be handled in jsCallMember
	obj["__js_cwd_method"] = NewString("__js_cwd__")
	obj["__js_exit_method"] = NewString("__js_exit__")

	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 10)

	// Mark this as the process object for special handling
	vm.jsProcessObjectID = objID

	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateEnvObject creates a JS object representation of process.env
// mapping environment variables to their values.
func (vm *VM) jsCreateEnvObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value)

	// Add all environment variables
	for _, envVar := range os.Environ() {
		parts := strings.SplitN(envVar, "=", 2)
		if len(parts) == 2 {
			key := parts[0]
			value := parts[1]
			obj[key] = NewString(value)
		}
	}

	obj["__js_type"] = NewString("process.env")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, len(obj))

	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateArgvArray creates a JS array of command-line arguments.
// In CLI mode, this maps to os.Args.
// In HTTP/FastCGI server mode, this is a minimal array with just the executable name.
func (vm *VM) jsCreateArgvArray() Value {
	var args []Value

	// First element is the executable path
	if len(os.Args) > 0 {
		args = append(args, NewString(os.Args[0]))
	} else {
		args = append(args, NewString("axonasp"))
	}

	// Append remaining arguments (only if running in CLI mode)
	if len(os.Args) > 1 {
		for _, arg := range os.Args[1:] {
			args = append(args, NewString(arg))
		}
	}

	vbArray := NewVBArrayFromValues(0, args)
	return ValueFromVBArray(vbArray)
}

// jsCallProcessMethod handles special process object methods like cwd() and exit().
// Returns (value, handled bool) - if handled=true, the value is the result.
func (vm *VM) jsCallProcessMethod(methodName string, args []Value) (Value, bool) {
	lower := strings.ToLower(methodName)

	switch lower {
	case "cwd":
		// process.cwd() - return current working directory
		cwd, err := os.Getwd()
		if err != nil {
			vm.jsThrowTypeError("Failed to get current working directory")
			return Value{Type: VTJSUndefined}, true
		}
		return NewString(cwd), true

	case "exit":
		// process.exit(code) - terminate the process
		code := 0
		if len(args) > 0 {
			code = int(vm.jsToNumber(args[0]).Flt)
		}

		// In CLI mode, use os.Exit directly
		// In HTTP/FastCGI mode, we need to safely terminate the JScript execution
		if vm.host != nil && vm.host.Request() != nil && vm.host.Request().ServerVars.Get("AXONASP_CLI_TUI") != "1" {
			// Server mode: safely terminate via panic/recover mechanism
			vm.jsThrowReferenceError("process.exit(" + fmt.Sprintf("%d", code) + ")")
			return Value{Type: VTJSUndefined}, true
		}

		// CLI mode: exit directly
		os.Exit(code)
		return Value{Type: VTJSUndefined}, true

	case "env":
		// process.env is accessed as a property, not a method
		// This shouldn't be called as a method
		return Value{Type: VTJSUndefined}, true

	case "argv":
		// process.argv is accessed as a property, not a method
		// This shouldn't be called as a method
		return Value{Type: VTJSUndefined}, true
	}

	return Value{Type: VTJSUndefined}, false
}
