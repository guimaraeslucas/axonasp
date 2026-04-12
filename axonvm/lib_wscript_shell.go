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
	"io"
	"os"
	"os/exec"
	"runtime"
	"strings"
	"sync"
	"time"
)

type WScriptShell struct {
	vm *VM
}

type WScriptExecObject struct {
	vm           *VM
	cmd          *exec.Cmd
	stdoutPipe   io.ReadCloser
	stderrPipe   io.ReadCloser
	stdinPipe    io.WriteCloser
	stdoutStream *ProcessTextStream
	stderrStream *ProcessTextStream
	status       int // 0 = Running, 1 = Done
	exitCode     int
	processID    int
	mu           sync.Mutex
	finished     bool
}

type ProcessTextStream struct {
	vm            *VM
	buffer        *bytes.Buffer
	pipe          io.ReadCloser
	atEndOfStream bool
	readingDone   bool
	mu            sync.Mutex
	closed        bool
	isStdout      bool
	isStderr      bool
}

// newWScriptShellObject instantiates the WScript.Shell library
func (vm *VM) newWScriptShellObject() Value {
	obj := &WScriptShell{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.wscriptShellItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet acts as a getter.
func (ws *WScriptShell) DispatchPropertyGet(propertyName string) Value {
	return ws.DispatchMethod(propertyName, nil)
}

// DispatchMethod provides O(1) string matching resolution.
func (ws *WScriptShell) DispatchMethod(methodName string, args []Value) Value {
	method := strings.ToLower(methodName)

	switch method {
	case "run":
		if len(args) < 1 {
			return NewInteger(-1)
		}
		command := args[0].String()
		if command == "" {
			return NewInteger(-1)
		}

		waitOnReturn := true
		if len(args) > 2 {
			waitOnReturn = args[2].Num != 0
		}

		var cmd *exec.Cmd
		if runtime.GOOS == "windows" {
			cmd = exec.Command("cmd.exe", "/c", command)
		} else {
			cmd = exec.Command("sh", "-c", command)
		}

		if waitOnReturn {
			err := cmd.Run()
			if err != nil {
				if exitErr, ok := err.(*exec.ExitError); ok {
					return NewInteger(int64(exitErr.ExitCode()))
				}
				return NewInteger(-1)
			}
			return NewInteger(0)
		} else {
			err := cmd.Start()
			if err != nil {
				return NewInteger(-1)
			}
			return NewInteger(0)
		}

	case "exec":
		if len(args) < 1 {
			return NewEmpty()
		}
		command := args[0].String()
		if command == "" {
			return NewEmpty()
		}

		var cmd *exec.Cmd
		if runtime.GOOS == "windows" {
			cmd = exec.Command("cmd.exe", "/c", command)
		} else {
			cmd = exec.Command("sh", "-c", command)
		}

		stdoutPipe, err := cmd.StdoutPipe()
		if err != nil {
			return NewEmpty()
		}

		stderrPipe, err := cmd.StderrPipe()
		if err != nil {
			return NewEmpty()
		}

		stdinPipe, err := cmd.StdinPipe()
		if err != nil {
			return NewEmpty()
		}

		err = cmd.Start()
		if err != nil {
			return NewEmpty()
		}

		stdoutBuffer := &bytes.Buffer{}
		stderrBuffer := &bytes.Buffer{}

		execObj := &WScriptExecObject{
			vm:         ws.vm,
			cmd:        cmd,
			stdoutPipe: stdoutPipe,
			stderrPipe: stderrPipe,
			stdinPipe:  stdinPipe,
			status:     0,
			exitCode:   -1,
			processID:  cmd.Process.Pid,
			finished:   false,
		}

		execObj.stdoutStream = &ProcessTextStream{
			vm:            ws.vm,
			buffer:        stdoutBuffer,
			pipe:          stdoutPipe,
			atEndOfStream: false,
			isStdout:      true,
		}

		execObj.stderrStream = &ProcessTextStream{
			vm:            ws.vm,
			buffer:        stderrBuffer,
			pipe:          stderrPipe,
			atEndOfStream: false,
			isStderr:      true,
		}

		go func() {
			io.Copy(stdoutBuffer, stdoutPipe)
			execObj.stdoutStream.mu.Lock()
			execObj.stdoutStream.readingDone = true
			execObj.stdoutStream.mu.Unlock()
		}()

		go func() {
			io.Copy(stderrBuffer, stderrPipe)
			execObj.stderrStream.mu.Lock()
			execObj.stderrStream.readingDone = true
			execObj.stderrStream.mu.Unlock()
		}()

		go func() {
			err := execObj.cmd.Wait()
			execObj.mu.Lock()
			defer execObj.mu.Unlock()
			execObj.status = 1
			if err != nil {
				if exitErr, ok := err.(*exec.ExitError); ok {
					execObj.exitCode = exitErr.ExitCode()
				} else {
					execObj.exitCode = -1
				}
			} else {
				execObj.exitCode = 0
			}
			execObj.finished = true
			execObj.stdoutStream.mu.Lock()
			execObj.stdoutStream.atEndOfStream = true
			execObj.stdoutStream.mu.Unlock()
			execObj.stderrStream.mu.Lock()
			execObj.stderrStream.atEndOfStream = true
			execObj.stderrStream.mu.Unlock()
		}()

		id := ws.vm.nextDynamicNativeID
		ws.vm.nextDynamicNativeID++
		ws.vm.wscriptExecItems[id] = execObj
		return Value{Type: VTNativeObject, Num: id}

	case "createobject":
		if len(args) < 1 {
			return NewEmpty()
		}
		// Delegate to vm
		return ws.vm.dispatchNativeCall(nativeObjectServer, "CreateObject", args)

	case "getenv", "environmentvariables":
		if len(args) < 1 {
			return NewString("")
		}
		return NewString(os.Getenv(args[0].String()))

	case "expandenvironmentstrings":
		if len(args) < 1 {
			return NewString("")
		}
		return NewString(expandEnvironmentStrings(args[0].String()))
	}
	return NewEmpty()
}

// expandEnvironmentStrings expands Windows-style %NAME% environment placeholders.
// Unknown placeholders are preserved exactly as written for compatibility.
func expandEnvironmentStrings(input string) string {
	if input == "" {
		return ""
	}

	var b strings.Builder
	b.Grow(len(input))

	for i := 0; i < len(input); {
		if input[i] != '%' {
			b.WriteByte(input[i])
			i++
			continue
		}

		start := i
		i++
		nameStart := i
		for i < len(input) && input[i] != '%' {
			i++
		}
		if i >= len(input) {
			// No closing '%' found; preserve trailing content unchanged.
			b.WriteString(input[start:])
			break
		}

		name := input[nameStart:i]
		if name == "" {
			// Preserve %% literally.
			b.WriteString("%%")
			i++
			continue
		}

		if value, ok := os.LookupEnv(name); ok {
			b.WriteString(value)
		} else {
			// Keep unknown variables as %NAME% to match WScript.Shell behavior.
			b.WriteByte('%')
			b.WriteString(name)
			b.WriteByte('%')
		}
		i++
	}

	return b.String()
}

func (we *WScriptExecObject) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(propertyName) {
	case "status":
		we.mu.Lock()
		defer we.mu.Unlock()
		return NewInteger(int64(we.status))
	case "exitcode":
		we.mu.Lock()
		defer we.mu.Unlock()
		if !we.finished {
			return NewInteger(-1)
		}
		return NewInteger(int64(we.exitCode))
	case "processid", "pid":
		we.mu.Lock()
		defer we.mu.Unlock()
		return NewInteger(int64(we.processID))
	case "stdout":
		id := we.vm.nextDynamicNativeID
		we.vm.nextDynamicNativeID++
		we.vm.wscriptProcessStreamItems[id] = we.stdoutStream
		return Value{Type: VTNativeObject, Num: id}
	case "stderr":
		id := we.vm.nextDynamicNativeID
		we.vm.nextDynamicNativeID++
		we.vm.wscriptProcessStreamItems[id] = we.stderrStream
		return Value{Type: VTNativeObject, Num: id}
	}
	return we.DispatchMethod(propertyName, nil)
}

func (we *WScriptExecObject) DispatchMethod(methodName string, args []Value) Value {
	switch strings.ToLower(methodName) {
	case "waituntildone":
		timeout := 0
		if len(args) > 0 {
			timeout = int(we.vm.asInt(args[0]))
		}
		totalWait := time.Duration(timeout) * time.Millisecond
		startTime := time.Now()
		for {
			we.mu.Lock()
			if we.finished {
				we.mu.Unlock()
				return NewBool(true)
			}
			we.mu.Unlock()
			if timeout > 0 && time.Since(startTime) > totalWait {
				return NewBool(false)
			}
			time.Sleep(100 * time.Millisecond)
		}
	case "terminate":
		we.mu.Lock()
		defer we.mu.Unlock()
		if we.cmd != nil && we.cmd.Process != nil {
			we.cmd.Process.Kill()
		}
		return NewEmpty()
	}
	return NewEmpty()
}

func (ts *ProcessTextStream) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(propertyName) {
	case "atendofstream":
		ts.mu.Lock()
		defer ts.mu.Unlock()
		return NewBool(ts.atEndOfStream)
	case "line":
		return NewInteger(0)
	}
	return ts.DispatchMethod(propertyName, nil)
}

func (ts *ProcessTextStream) DispatchMethod(methodName string, args []Value) Value {
	switch strings.ToLower(methodName) {
	case "read":
		numChars := 1
		if len(args) > 0 {
			numChars = int(ts.vm.asInt(args[0]))
		}
		ts.mu.Lock()
		defer ts.mu.Unlock()
		if ts.closed || ts.atEndOfStream {
			return NewString("")
		}
		buffer := make([]byte, numChars)
		n, err := ts.pipe.Read(buffer)
		if err != nil && err != io.EOF {
			return NewString("")
		}
		if n == 0 {
			ts.atEndOfStream = true
			return NewString("")
		}
		return NewString(string(buffer[:n]))
	case "readline":
		ts.mu.Lock()
		defer ts.mu.Unlock()
		if ts.closed || ts.atEndOfStream {
			return NewString("")
		}
		buffer := make([]byte, 4096)
		n, err := ts.pipe.Read(buffer)
		if err != nil && err != io.EOF {
			return NewString("")
		}
		if n == 0 {
			ts.atEndOfStream = true
			return NewString("")
		}
		content := string(buffer[:n])
		lines := strings.Split(content, "\n")
		return NewString(strings.TrimSuffix(lines[0], "\r"))
	case "readall":
		for {
			ts.mu.Lock()
			if ts.closed || ts.readingDone {
				output := ts.buffer.String()
				ts.buffer.Reset()
				ts.atEndOfStream = true
				ts.mu.Unlock()
				return NewString(output)
			}
			ts.mu.Unlock()
			time.Sleep(10 * time.Millisecond)
		}
	case "close":
		ts.mu.Lock()
		defer ts.mu.Unlock()
		if !ts.closed {
			ts.pipe.Close()
			ts.closed = true
		}
		return NewEmpty()
	}
	return NewEmpty()
}
