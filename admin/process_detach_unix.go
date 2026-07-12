//go:build unix && !wasm

package main

import (
	"os/exec"
	"syscall"
)

// configureDetachedProcess ensures child executables are not bound to axonadmin lifecycle.
func configureDetachedProcess(cmd *exec.Cmd) {
	cmd.SysProcAttr = &syscall.SysProcAttr{
		Setsid:  true,
		Setpgid: true,
	}
}

// configureVisibleConsoleProcess keeps unix child configuration detached when no native console API exists.
func configureVisibleConsoleProcess(cmd *exec.Cmd) {
	configureDetachedProcess(cmd)
}
