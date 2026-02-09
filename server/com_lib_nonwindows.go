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
package server

import "fmt"

// NewCOMObject returns an error on non-Windows platforms.
func NewCOMObject(progID string) (*COMObject, error) {
	return nil, fmt.Errorf("COM is not supported on this platform")
}

// COMObject is a stub on non-Windows platforms.
type COMObject struct{}

// GetProperty returns nil on non-Windows platforms.
func (c *COMObject) GetProperty(name string) interface{} { return nil }

// SetProperty returns an error on non-Windows platforms.
func (c *COMObject) SetProperty(name string, value interface{}) error {
	return fmt.Errorf("COM is not supported on this platform")
}

// CallMethod returns an error on non-Windows platforms.
func (c *COMObject) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return nil, fmt.Errorf("COM is not supported on this platform")
}

// Enumerate returns an empty list on non-Windows platforms.
func (c *COMObject) Enumerate() ([]interface{}, error) { return []interface{}{}, nil }

// No-op on non-Windows platforms
func (c *COMObject) release() {
	// No-op on non-Windows platforms
}
