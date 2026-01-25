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

import (
	"crypto/rand"
	"fmt"
	"strings"

	"golang.org/x/crypto/bcrypt"
)

// G3CRYPTO implements Component interface for Crypto operations
type G3CRYPTO struct {
	ctx *ExecutionContext
}

func (c *G3CRYPTO) GetProperty(name string) interface{} {
	return nil
}

func (c *G3CRYPTO) SetProperty(name string, value interface{}) {}

func (c *G3CRYPTO) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)

	getStr := func(i int) string {
		if i >= len(args) {
			return ""
		}
		return fmt.Sprintf("%v", args[i])
	}

	switch method {
	case "uuid":
		return c.UUID()

	case "hashpassword", "hash":
		return c.HashPassword(getStr(0))

	case "verifypassword", "verify":
		pass := getStr(0)
		hash := getStr(1)
		return c.VerifyPassword(pass, hash)
	}

	return nil
}

func (c *G3CRYPTO) UUID() string {
	b := make([]byte, 16)
	_, err := rand.Read(b)
	if err != nil {
		return ""
	}
	b[6] = (b[6] & 0x0f) | 0x40
	b[8] = (b[8] & 0x3f) | 0x80
	return fmt.Sprintf("%x-%x-%x-%x-%x", b[0:4], b[4:6], b[6:8], b[8:10], b[10:])
}

func (c *G3CRYPTO) HashPassword(password string) string {
	bytes, err := bcrypt.GenerateFromPassword([]byte(password), bcrypt.DefaultCost)
	if err != nil {
		return ""
	}
	return string(bytes)
}

func (c *G3CRYPTO) VerifyPassword(password, hash string) bool {
	err := bcrypt.CompareHashAndPassword([]byte(hash), []byte(password))
	return err == nil
}

// CryptoHelper for legacy support
func CryptoHelper(method string, args []string, ctx *ExecutionContext) interface{} {
	lib := &G3CRYPTO{ctx: ctx}
	var ifaceArgs []interface{}
	for _, a := range args {
		ifaceArgs = append(ifaceArgs, EvaluateExpression(a, ctx))
	}
	return lib.CallMethod(method, ifaceArgs)
}
