/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas GuimarÃ£es - G3pix Ltda
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
	"io"
	"os"
	"os/exec"
	"runtime"
	"strings"
	"sync"
	"time"
)

// WScriptShell implements the WScript.Shell COM object
type WScriptShell struct {
	ctx *ExecutionContext
}

// WScriptExecObject represents the object returned by Shell.Exec()
type WScriptExecObject struct {
	cmd           *exec.Cmd
	stdoutPipe    io.ReadCloser
	stderrPipe    io.ReadCloser
	stdinPipe     io.WriteCloser
	stdoutStream  *ProcessTextStream
	stderrStream  *ProcessTextStream
	status        int // 0 = Running, 1 = Done
	exitCode      int
	processID     int
	mu            sync.Mutex
	finished      bool
}

// ProcessTextStream represents a text stream for reading/writing from process I/O
type ProcessTextStream struct {
	buffer        *bytes.Buffer
	pipe          io.ReadCloser
	atEndOfStream bool
	readingDone   bool // Set to true when background reader goroutine finishes
	mu            sync.Mutex
	closed        bool
	isStdout      bool
	isStderr      bool
}

// NewWScriptShell creates a new WScript.Shell object
func NewWScriptShell(ctx *ExecutionContext) *WScriptShell {
	return &WScriptShell{
		ctx: ctx,
	}
}

// GetProperty implements the Component interface
func (ws *WScriptShell) GetProperty(name string) interface{} {
	return nil
}

// SetProperty implements the Component interface
func (ws *WScriptShell) SetProperty(name string, value interface{}) {}

// CallMethod implements the Component interface
func (ws *WScriptShell) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)

	switch method {
	case "exec":
		return ws.Exec(args...)
	case "run":
		return ws.Run(args...)
	case "createobject":
		return ws.CreateObject(args...)
	case "getenv":
		return ws.GetEnv(args...)
	default:
		return nil
	}
}

// Run method: Shell.Run(strCommand, [intWindowStyle], [bWaitOnReturn])
// Returns: Exit code of the command (0 on success)
func (ws *WScriptShell) Run(args ...interface{}) interface{} {
	if len(args) < 1 {
		return nil
	}

	command := toString(args[0])
	if command == "" {
		return -1
	}

	// Window style (default 1 = normal)
	// 0 = hidden, 1 = normal, 2 = minimized, 3 = maximized, etc
	windowStyle := 1
	if len(args) > 1 {
		windowStyle = toInt(args[1])
	}

	// Wait on return (default true)
	waitOnReturn := true
	if len(args) > 2 {
		waitOnReturn = (toInt(args[2]) != 0)
	}

	var cmd *exec.Cmd

	// Build command based on OS
	if runtime.GOOS == "windows" {
		// Windows: use cmd.exe /c
		cmd = exec.Command("cmd.exe", "/c", command)
	} else {
		// Unix-like: use sh -c
		cmd = exec.Command("sh", "-c", command)
	}

	// Window style is ignored on non-Windows systems
	_ = windowStyle

	if waitOnReturn {
		// Synchronous execution
		err := cmd.Run()
		if err != nil {
			if exitErr, ok := err.(*exec.ExitError); ok {
				return exitErr.ExitCode()
			}
			return -1
		}
		return 0
	} else {
		// Asynchronous execution - fire and forget
		err := cmd.Start()
		if err != nil {
			return -1
		}
		// Return immediately with 0
		return 0
	}
}

// Exec method: Shell.Exec(strCommand)
// Returns: WScriptExecObject with access to StdOut, StdErr, StdIn, Status, and ProcessID
func (ws *WScriptShell) Exec(args ...interface{}) interface{} {
	if len(args) < 1 {
		return nil
	}

	command := toString(args[0])
	if command == "" {
		return nil
	}

	var cmd *exec.Cmd

	// Build command based on OS
	if runtime.GOOS == "windows" {
		// Windows: use cmd.exe /c
		cmd = exec.Command("cmd.exe", "/c", command)
	} else {
		// Unix-like: use sh -c
		cmd = exec.Command("sh", "-c", command)
	}

	// Create pipes for stdin, stdout, stderr
	stdoutPipe, err := cmd.StdoutPipe()
	if err != nil {
		return nil
	}

	stderrPipe, err := cmd.StderrPipe()
	if err != nil {
		return nil
	}

	stdinPipe, err := cmd.StdinPipe()
	if err != nil {
		return nil
	}

	// Start the command
	err = cmd.Start()
	if err != nil {
		return nil
	}

	// Create output buffers that will be filled by reading goroutines
	stdoutBuffer := &bytes.Buffer{}
	stderrBuffer := &bytes.Buffer{}

	// Create the WScriptExecObject
	execObj := &WScriptExecObject{
		cmd:        cmd,
		stdoutPipe: stdoutPipe,
		stderrPipe: stderrPipe,
		stdinPipe:  stdinPipe,
		status:     0, // Running
		exitCode:   -1,
		processID:  cmd.Process.Pid,
		finished:   false,
	}

	// Create TextStream objects for output
	execObj.stdoutStream = &ProcessTextStream{
		buffer:        stdoutBuffer,
		pipe:          stdoutPipe,
		atEndOfStream: false,
		isStdout:      true,
	}

	execObj.stderrStream = &ProcessTextStream{
		buffer:        stderrBuffer,
		pipe:          stderrPipe,
		atEndOfStream: false,
		isStderr:      true,
	}

	// Start goroutines to read stdout and stderr BEFORE waiting for command
	// This ensures pipes are read while command is running
	go func() {
		io.Copy(stdoutBuffer, stdoutPipe)
		// Mark reading as done
		execObj.stdoutStream.mu.Lock()
		execObj.stdoutStream.readingDone = true
		execObj.stdoutStream.mu.Unlock()
	}()

	go func() {
		io.Copy(stderrBuffer, stderrPipe)
		// Mark reading as done
		execObj.stderrStream.mu.Lock()
		execObj.stderrStream.readingDone = true
		execObj.stderrStream.mu.Unlock()
	}()

	// Start a goroutine to wait for command completion and update status
	go func() {
		err := execObj.cmd.Wait()

		// Update status and exit code
		execObj.mu.Lock()
		defer execObj.mu.Unlock()
		
		execObj.status = 1 // Done

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
		
		// Mark stream as at end since we've finished reading
		execObj.stdoutStream.mu.Lock()
		execObj.stdoutStream.atEndOfStream = true
		execObj.stdoutStream.mu.Unlock()

		execObj.stderrStream.mu.Lock()
		execObj.stderrStream.atEndOfStream = true
		execObj.stderrStream.mu.Unlock()
	}()

	return execObj
}

// CreateObject method for backward compatibility
func (ws *WScriptShell) CreateObject(args ...interface{}) interface{} {
	if len(args) < 1 {
		return nil
	}
	// Delegate to the executor's CreateObject
	if ws.ctx != nil {
		// This would need integration with executor
		// For now, return nil
	}
	return nil
}

// GetEnv method: Shell.EnvironmentVariables() or Shell.GetEnv(name)
// Note: In classic ASP, EnvironmentVariables is a collection, we'll support GetEnv
func (ws *WScriptShell) GetEnv(args ...interface{}) interface{} {
	if len(args) < 1 {
		return ""
	}

	envName := toString(args[0])
	if envName == "" {
		return ""
	}

	return os.Getenv(envName)
}

// WScriptExecObject Methods and Properties

// GetProperty implements the Component interface for WScriptExecObject
func (we *WScriptExecObject) GetProperty(name string) interface{} {
	name = strings.ToLower(name)

	switch name {
	case "status":
		we.mu.Lock()
		defer we.mu.Unlock()
		return we.status

	case "exitcode":
		we.mu.Lock()
		defer we.mu.Unlock()
		if !we.finished {
			return -1
		}
		return we.exitCode

	case "processid", "pid":
		we.mu.Lock()
		defer we.mu.Unlock()
		return we.processID

	case "stdout":
		return we.stdoutStream

	case "stderr":
		return we.stderrStream

	case "stdin":
		return &StdInStream{
			pipe: we.stdinPipe,
		}

	default:
		return nil
	}
}

// SetProperty implements the Component interface
func (we *WScriptExecObject) SetProperty(name string, value interface{}) {}

// CallMethod implements the Component interface
func (we *WScriptExecObject) CallMethod(name string, args ...interface{}) interface{} {
	name = strings.ToLower(name)
	
	switch name {
	case "waituntildone":
		timeout := 0  // Default: no timeout (wait indefinitely)
		if len(args) > 0 {
			timeout = toInt(args[0])
		}
		return we.WaitUntilDone(timeout)
	
	case "terminate":
		we.Terminate()
		return nil
	
	default:
		return nil
	}
}

// ProcessTextStream Methods and Properties

// GetProperty implements the Component interface for ProcessTextStream
func (ts *ProcessTextStream) GetProperty(name string) interface{} {
	name = strings.ToLower(name)

	switch name {
	case "atendofstream":
		ts.mu.Lock()
		defer ts.mu.Unlock()
		return ts.atEndOfStream

	case "line":
		// Current line number would require tracking
		return 0

	default:
		return nil
	}
}

// SetProperty implements the Component interface
func (ts *ProcessTextStream) SetProperty(name string, value interface{}) {}

// CallMethod implements the Component interface
func (ts *ProcessTextStream) CallMethod(name string, args ...interface{}) interface{} {
	name = strings.ToLower(name)

	switch name {
	case "read":
		numChars := 1
		if len(args) > 0 {
			numChars = toInt(args[0])
		}
		return ts.Read(numChars)

	case "readline":
		return ts.ReadLine()

	case "readall":
		return ts.ReadAll()

	case "close":
		ts.Close()
		return nil

	default:
		return nil
	}
}

// Read reads specified number of characters from the stream
func (ts *ProcessTextStream) Read(numChars int) string {
	ts.mu.Lock()
	defer ts.mu.Unlock()

	if ts.closed || ts.atEndOfStream {
		return ""
	}

	buffer := make([]byte, numChars)
	n, err := ts.pipe.Read(buffer)

	if err != nil && err != io.EOF {
		return ""
	}

	if n == 0 {
		ts.atEndOfStream = true
		return ""
	}

	return string(buffer[:n])
}

// ReadLine reads one line from the stream
func (ts *ProcessTextStream) ReadLine() string {
	ts.mu.Lock()
	defer ts.mu.Unlock()

	if ts.closed || ts.atEndOfStream {
		return ""
	}

	buffer := make([]byte, 4096)
	n, err := ts.pipe.Read(buffer)

	if err != nil && err != io.EOF {
		return ""
	}

	if n == 0 {
		ts.atEndOfStream = true
		return ""
	}

	// Split by newline and return first line
	content := string(buffer[:n])
	lines := strings.Split(content, "\n")

	return strings.TrimSuffix(lines[0], "\r")
}

// ReadAll reads all remaining content from the stream
func (ts *ProcessTextStream) ReadAll() string {
	// Wait for the background reader goroutine to finish reading from pipe
	for {
		ts.mu.Lock()
		if ts.closed || ts.readingDone {
			output := ts.buffer.String()
			ts.buffer.Reset()
			ts.atEndOfStream = true
			ts.mu.Unlock()
			return output
		}
		ts.mu.Unlock()

		// Small sleep to avoid busy waiting
		time.Sleep(10 * time.Millisecond)
	}
}

// Close closes the stream
func (ts *ProcessTextStream) Close() {
	ts.mu.Lock()
	defer ts.mu.Unlock()

	if !ts.closed {
		ts.pipe.Close()
		ts.closed = true
	}
}

// StdInStream for input handling
type StdInStream struct {
	pipe io.WriteCloser
	mu   sync.Mutex
}

// GetProperty for StdInStream
func (sis *StdInStream) GetProperty(name string) interface{} {
	return nil
}

// SetProperty for StdInStream
func (sis *StdInStream) SetProperty(name string, value interface{}) {}

// CallMethod for StdInStream
func (sis *StdInStream) CallMethod(name string, args ...interface{}) interface{} {
	name = strings.ToLower(name)

	switch name {
	case "close":
		sis.Close()
		return nil

	default:
		return nil
	}
}

// Write writes to stdin
func (sis *StdInStream) Write(text string) {
	sis.mu.Lock()
	defer sis.mu.Unlock()

	if sis.pipe != nil {
		sis.pipe.Write([]byte(text))
	}
}

// WriteLine writes to stdin with newline
func (sis *StdInStream) WriteLine(text string) {
	sis.mu.Lock()
	defer sis.mu.Unlock()

	if sis.pipe != nil {
		sis.pipe.Write([]byte(text + "\n"))
	}
}

// Close closes stdin
func (sis *StdInStream) Close() {
	sis.mu.Lock()
	defer sis.mu.Unlock()

	if sis.pipe != nil {
		sis.pipe.Close()
		sis.pipe = nil
	}
}

// WaitUntilDone waits for the process to complete
func (we *WScriptExecObject) WaitUntilDone(timeout int) bool {
	totalWait := time.Duration(timeout) * time.Millisecond
	startTime := time.Now()

	for {
		we.mu.Lock()
		if we.finished {
			we.mu.Unlock()
			return true
		}
		we.mu.Unlock()

		if timeout > 0 && time.Since(startTime) > totalWait {
			return false
		}

		time.Sleep(100 * time.Millisecond)
	}
}

// Terminate kills the process
func (we *WScriptExecObject) Terminate() {
	we.mu.Lock()
	defer we.mu.Unlock()

	if we.cmd != nil && we.cmd.Process != nil {
		we.cmd.Process.Kill()
	}
}



