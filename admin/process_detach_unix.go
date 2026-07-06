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
