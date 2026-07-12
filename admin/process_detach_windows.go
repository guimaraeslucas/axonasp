//go:build windows && !wasm

package main

import (
	"os/exec"
	"syscall"
)

const (
	createNewProcessGroup = 0x00000200
	detachedProcess       = 0x00000008
	createNoWindow        = 0x08000000
	createNewConsole      = 0x00000010
)

// configureDetachedProcess ensures child executables are not bound to axonadmin lifecycle.
func configureDetachedProcess(cmd *exec.Cmd) {
	cmd.SysProcAttr = &syscall.SysProcAttr{
		CreationFlags: createNewProcessGroup | detachedProcess | createNoWindow,
	}
}

// configureVisibleConsoleProcess starts a process attached to a new visible console window.
func configureVisibleConsoleProcess(cmd *exec.Cmd) {
	cmd.SysProcAttr = &syscall.SysProcAttr{
		CreationFlags: createNewProcessGroup | createNewConsole,
	}
}
