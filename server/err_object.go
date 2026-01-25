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

import "strings"

// ErrObject models the classic ASP Err intrinsic object.
type ErrObject struct {
	Number      int
	Description string
	Source      string
}

// NewErrObject builds a clean ErrObject instance.
func NewErrObject() *ErrObject {
	return &ErrObject{}
}

// Clear resets the current error state.
func (e *ErrObject) Clear() {
	e.Number = 0
	e.Description = ""
	e.Source = ""
}

// Raise sets the error fields using the provided values.
func (e *ErrObject) Raise(number int, source string, description string) {
	e.Number = number
	e.Source = source
	e.Description = description
}

// SetError maps a Go error into the Err object (used by On Error Resume Next flow).
func (e *ErrObject) SetError(err error) {
	if err == nil {
		e.Clear()
		return
	}
	e.Number = -1
	e.Source = "Runtime"
	e.Description = err.Error()
}

// GetName returns the ASP object name.
func (e *ErrObject) GetName() string {
	return "Err"
}

// GetProperty exposes Err members to VBScript.
func (e *ErrObject) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "number":
		return e.Number
	case "description":
		return e.Description
	case "source":
		return e.Source
	default:
		return nil
	}
}

// SetProperty allows VBScript to assign Err members.
func (e *ErrObject) SetProperty(name string, value interface{}) error {
	switch strings.ToLower(name) {
	case "number":
		e.Number = toInt(value)
	case "description":
		e.Description = toString(value)
	case "source":
		e.Source = toString(value)
	}
	return nil
}

// CallMethod supports Err.Clear and Err.Raise.
func (e *ErrObject) CallMethod(name string, args ...interface{}) (interface{}, error) {
	switch strings.ToLower(name) {
	case "clear":
		e.Clear()
		return nil, nil
	case "raise":
		var number int
		var source string
		var description string
		if len(args) > 0 {
			number = toInt(args[0])
		}
		if len(args) > 1 {
			source = toString(args[1])
		}
		if len(args) > 2 {
			description = toString(args[2])
		}
		e.Raise(number, source, description)
		return nil, nil
	default:
		return nil, nil
	}
}
