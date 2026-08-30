/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães (@guimaraeslucas)
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

// VBScript MsgBox / InputBox OS-native dialog implementation for AxonHTA.
//
// WHY PHASE 1 (JS INJECTION) IS ARCHITECTURALLY IMPOSSIBLE
// ─────────────────────────────────────────────────────────
// axonhta's HTTP handler runs the VBScript VM in a goroutine and blocks on
// a channel (`done <- vmResult`), waiting for the script to finish before
// flushing the HTTP response. For MsgBox/InputBox to "call the browser" and
// wait for the user's answer, the sequence would be:
//
//   Go VM goroutine → inject JS alert/prompt into already-sent response
//                   → browser executes modal
//                   → browser POSTs result back to Go
//                   → Go unblocks VM goroutine
//
// The fatal flaw is step 1: the HTTP response has NOT been sent yet (the
// htmlInjectWriter buffers everything until FinalFlush() runs after the
// handler returns). The browser has no page to show the modal in. Even if
// the response were flushed early, the VM goroutine would need to block on
// a second HTTP request — a deadlock because the Go HTTP server's per-request
// goroutine is the one that is *also* waiting for the VM to finish. A
// separate polling endpoint would require multi-request sequencing that is
// fundamentally asynchronous and cannot be made synchronous without
// cooperative VBScript suspension (which the single-pass bytecode VM does
// not support).
//
// PROVIDER HOOK MECHANISM
// ────────────────────────
// axonvm.DialogMsgBoxProvider and axonvm.DialogInputBoxProvider are
// package-level function variables (default nil). The vbsMsgBox / vbsInputBox
// stubs in axonvm/builtins.go check these at call time; when non-nil they
// delegate. Setting them in init() here is safe: Go guarantees all init()
// functions complete before main() runs, so the providers are in place
// before the HTTP server accepts its first request.
//
// WHY ncruces/zenity FOR PHASE 2
// ──────────────────────────────
// github.com/ncruces/zenity provides blocking, synchronous OS dialogs:
//   - Windows: pure-Go COM-based TaskDialog / VBScript InputBox (zero CGO)
//   - macOS:   pure-Go Cocoa NSAlert / NSTextField sheet
//   - Linux:   zenity/kdialog subprocess (pure Go, no CGO)
//
// The call blocks the VM goroutine while the dialog is open and returns
// the user's response directly — matching VBScript semantics exactly.

import (
	"strings"

	"g3pix.com.br/axonasp/v2/axonvm"
	"github.com/ncruces/zenity"
)

// init sets the axonvm dialog providers before the HTTP server starts.
// This is the correct point: axonvm's own init() has already run (imported
// package init runs first), so the builtin registry is complete. We are
// not re-registering builtins — we are setting runtime dispatch hooks that
// vbsMsgBox/vbsInputBox check on every call.
func init() {
	axonvm.DialogMsgBoxProvider = htaMsgBox
	axonvm.DialogInputBoxProvider = htaInputBox
}

// htaMsgBox is the HTA-specific implementation of VBScript MsgBox().
// Blocks the calling goroutine and shows an OS-native modal dialog.
//
// Signature: MsgBox(prompt [, buttons [, title]])
//
// Button group constants (lower nibble of buttons):
//
//	0 = vbOKOnly              → OK
//	1 = vbOKCancel            → OK / Cancel
//	2 = vbAbortRetryIgnore    → mapped to Retry / Ignore
//	3 = vbYesNoCancel         → Yes / No (Cancel treated as No)
//	4 = vbYesNo               → Yes / No
//	5 = vbRetryCancel         → Retry / Cancel
//
// Return values follow VBScript constants:
//
//	1=vbOK 2=vbCancel 4=vbRetry 5=vbIgnore 6=vbYes 7=vbNo
func htaMsgBox(args []axonvm.Value) (axonvm.Value, error) {
	if len(args) < 1 {
		return axonvm.NewInteger(1), nil // vbOK
	}

	prompt := args[0].String()

	buttons := 0
	if len(args) >= 2 {
		buttons = int(args[1].Num)
	}

	dialogTitle := "AxonHTA"
	if len(args) >= 3 && args[2].String() != "" {
		dialogTitle = args[2].String()
	}

	btnGroup := buttons & 0xF

	switch btnGroup {
	case 1: // vbOKCancel
		err := zenity.Question(prompt,
			zenity.Title(dialogTitle),
			zenity.OKLabel("OK"),
			zenity.CancelLabel("Cancel"),
		)
		if err == nil {
			return axonvm.NewInteger(1), nil // vbOK
		}
		return axonvm.NewInteger(2), nil // vbCancel

	case 2: // vbAbortRetryIgnore → map to Retry / Ignore
		err := zenity.Question(prompt,
			zenity.Title(dialogTitle),
			zenity.OKLabel("Retry"),
			zenity.CancelLabel("Ignore"),
		)
		if err == nil {
			return axonvm.NewInteger(4), nil // vbRetry
		}
		return axonvm.NewInteger(5), nil // vbIgnore

	case 3: // vbYesNoCancel → treat window-close as No
		err := zenity.Question(prompt,
			zenity.Title(dialogTitle),
			zenity.OKLabel("Yes"),
			zenity.CancelLabel("No"),
		)
		if err == nil {
			return axonvm.NewInteger(6), nil // vbYes
		}
		return axonvm.NewInteger(7), nil // vbNo

	case 4: // vbYesNo
		err := zenity.Question(prompt,
			zenity.Title(dialogTitle),
			zenity.OKLabel("Yes"),
			zenity.CancelLabel("No"),
		)
		if err == nil {
			return axonvm.NewInteger(6), nil // vbYes
		}
		return axonvm.NewInteger(7), nil // vbNo

	case 5: // vbRetryCancel
		err := zenity.Question(prompt,
			zenity.Title(dialogTitle),
			zenity.OKLabel("Retry"),
			zenity.CancelLabel("Cancel"),
		)
		if err == nil {
			return axonvm.NewInteger(4), nil // vbRetry
		}
		return axonvm.NewInteger(2), nil // vbCancel

	default: // vbOKOnly (0) and any other value
		_ = zenity.Info(prompt, zenity.Title(dialogTitle))
		return axonvm.NewInteger(1), nil // vbOK
	}
}

// htaInputBox is the HTA-specific implementation of VBScript InputBox().
// Blocks the calling goroutine and shows an OS-native text-entry dialog.
//
// Signature: InputBox(prompt [, title [, default]])
//
// Returns the user's input string, or "" if the user cancelled.
func htaInputBox(args []axonvm.Value) (axonvm.Value, error) {
	if len(args) < 1 {
		return axonvm.NewString(""), nil
	}

	prompt := args[0].String()

	dialogTitle := "AxonHTA"
	if len(args) >= 2 && args[1].String() != "" {
		dialogTitle = args[1].String()
	}

	defaultValue := ""
	if len(args) >= 3 {
		defaultValue = args[2].String()
	}

	result, err := zenity.Entry(prompt,
		zenity.Title(dialogTitle),
		zenity.EntryText(defaultValue),
	)
	if err != nil {
		// User cancelled or closed — VBScript InputBox returns "" on cancel.
		return axonvm.NewString(""), nil
	}

	return axonvm.NewString(strings.TrimRight(result, "\r\n")), nil
}
