//go:build js && wasm

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
package main

import (
	"bytes"
	"syscall/js"

	"g3pix.com.br/axonasp/v2/axonvm"
	"g3pix.com.br/axonasp/v2/axonvm/asp"
)

var (
	sharedApplication = asp.NewApplication()
	sharedSession     = asp.NewSession()
)

// executeASP handles the invocation from JS: AxonASP.execute(code)
func executeASP(this js.Value, args []js.Value) interface{} {
	promiseConstructor := js.Global().Get("Promise")

	return promiseConstructor.New(js.FuncOf(func(this js.Value, promiseArgs []js.Value) interface{} {
		resolve := promiseArgs[0]
		reject := promiseArgs[1]

		go func() {
			if len(args) == 0 {
				reject.Invoke("Error: No code provided")
				return
			}
			code := args[0].String()

			compiler := axonvm.NewASPCompiler(code)
			compiler.SetSourceName("/wasm.asp")

			if err := compiler.Compile(); err != nil {
				reject.Invoke("Compile Error: " + err.Error())
				return
			}

			vm := axonvm.AcquireVMFromCompiler(compiler)
			defer vm.Release()

			var outBuf bytes.Buffer
			host := axonvm.NewMockHost()
			host.SetOutput(&outBuf)

			host.SetApplication(sharedApplication)
			host.SetSession(sharedSession)
			vm.SetHost(host)

			runErr := vm.Run()
			host.Response().Flush()
			host.Response().ReleaseBuffer()

			output := outBuf.String()
			if runErr != nil {
				output += "\nRuntime Error: " + runErr.Error()
			}

			resolve.Invoke(output)
		}()

		return nil
	}))
}

// dispatchLiveEvent handles AxonLive events originating from the browser DOM.
// JS usage: AxonASP.dispatchLiveEvent(sessionID, componentID, eventName, eventArgsJSON)
func dispatchLiveEvent(this js.Value, args []js.Value) interface{} {
	promiseConstructor := js.Global().Get("Promise")

	return promiseConstructor.New(js.FuncOf(func(this js.Value, promiseArgs []js.Value) interface{} {
		resolve := promiseArgs[0]
		reject := promiseArgs[1]

		go func() {
			if len(args) < 4 {
				reject.Invoke("Error: Invalid arguments to dispatchLiveEvent")
				return
			}

			// Extract event details
			sessionID := args[0].String()
			componentID := args[1].String()
			eventName := args[2].String()
			eventArgsJSON := args[3].String()

			println("❖ AxonASP Bridge: Event Received -", eventName, "on", componentID)

			// In WASM, since there's no server file system, we need to know the script code.
			var code string
			if len(args) >= 5 && args[4].Type() == js.TypeString {
				code = args[4].String()
			} else {
				reject.Invoke("Error: ASP source code must be provided as the 5th argument for WASM AxonLive")
				return
			}

			compiler := axonvm.NewASPCompiler(code)
			compiler.SetSourceName("/wasm.asp")

			if err := compiler.Compile(); err != nil {
				reject.Invoke("Compile Error: " + err.Error())
				return
			}

			vm := axonvm.AcquireVMFromCompiler(compiler)
			defer vm.Release()

			var outBuf bytes.Buffer
			host := axonvm.NewMockHost()
			host.SetOutput(&outBuf)

			// Setup AxonLive fake HTTP headers so initPage() detects an async request
			host.Request().ServerVars.Add("HTTP_X_G3AXONLIVE", "true")
			host.Request().ServerVars.Add("HTTP_X_G3AXONLIVE_SESSIONID", sessionID)
			host.Request().ServerVars.Add("HTTP_X_G3AXONLIVE_COMPONENTID", componentID)
			host.Request().ServerVars.Add("HTTP_X_G3AXONLIVE_EVENTNAME", eventName)
			host.Request().ServerVars.Add("HTTP_X_G3AXONLIVE_EVENTARGS", eventArgsJSON)

			host.SetApplication(sharedApplication)
			host.SetSession(sharedSession)
			vm.SetHost(host)

			runErr := vm.Run()
			host.Response().Flush()
			host.Response().ReleaseBuffer()

			output := outBuf.String()
			if runErr != nil {
				output += "\nRuntime Error: " + runErr.Error()
			}

			resolve.Invoke(output)
		}()

		return nil
	}))
}
func main() {
	js.Global().Set("AxonASP", js.ValueOf(map[string]interface{}{
		"execute":           js.FuncOf(executeASP),
		"dispatchLiveEvent": js.FuncOf(dispatchLiveEvent),
	}))

	c := make(chan struct{})
	println("G3pix ❖ AxonASP initialized")
	<-c
}
