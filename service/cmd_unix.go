//go:build !windows

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
	"os/exec"
	"syscall"
)

// execCmdWrapper wraps exec.Cmd for shared program state.
type execCmdWrapper struct {
	*exec.Cmd
}

// buildOSCommand creates a child command configured for Unix service execution.
func buildOSCommand(executablePath string, env []string) *exec.Cmd {
	cmd := exec.Command(executablePath)
	if len(env) > 0 {
		cmd.Env = env
	}
	return cmd
}

// stopOSCommand requests graceful termination on Unix before force kill fallback.
func stopOSCommand(cmd *exec.Cmd) error {
	if err := cmd.Process.Signal(syscall.SIGTERM); err != nil {
		return cmd.Process.Kill()
	}
	return nil
}
