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
	"fmt"
	"os"
	"strconv"
	"strings"

	"gopkg.in/gomail.v2"
)

type G3Mail struct {
	vm       *VM
	host     string
	port     int
	username string
	password string
	from     string
	fromName string
	to       []string
	cc       []string
	bcc      []string
	subject  string
	body     string
	isHTML   bool
}

// newG3MailObject instantiates the G3Mail custom functions library.
func (vm *VM) newG3MailObject() Value {
	obj := &G3Mail{
		vm:  vm,
		to:  make([]string, 0),
		cc:  make([]string, 0),
		bcc: make([]string, 0),
	}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3mailItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet acts as a getter.
func (m *G3Mail) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(propertyName) {
	case "host", "mailhost", "smtpserver":
		return NewString(m.host)
	case "port", "smtpserverport":
		return NewInteger(int64(m.port))
	case "username", "user", "authusername":
		return NewString(m.username)
	case "password", "pass", "authpassword":
		return NewString(m.password)
	case "from", "fromaddress":
		return NewString(m.from)
	case "fromname":
		return NewString(m.fromName)
	case "to":
		return NewString(strings.Join(m.to, ","))
	case "cc":
		return NewString(strings.Join(m.cc, ","))
	case "bcc":
		return NewString(strings.Join(m.bcc, ","))
	case "subject":
		return NewString(m.subject)
	case "body", "textbody", "htmlbody", "message":
		return NewString(m.body)
	case "ishtml":
		return NewBool(m.isHTML)
	case "bodyformat", "mailformat":
		if m.isHTML {
			return NewInteger(0)
		}
		return NewInteger(1)
	}
	return m.DispatchMethod(propertyName, nil)
}

// DispatchPropertySet acts as a setter.
func (m *G3Mail) DispatchPropertySet(propertyName string, args []Value) bool {
	if len(args) == 0 {
		return false
	}
	value := args[0]
	valueStr := m.vm.valueToString(value)

	switch strings.ToLower(propertyName) {
	case "host", "mailhost", "smtpserver":
		m.host = strings.TrimSpace(valueStr)
		return true
	case "port", "smtpserverport":
		m.port = int(m.vm.asInt(value))
		return true
	case "username", "user", "authusername":
		m.username = strings.TrimSpace(valueStr)
		return true
	case "password", "pass", "authpassword":
		m.password = valueStr
		return true
	case "from", "fromaddress":
		m.from = strings.TrimSpace(valueStr)
		return true
	case "fromname":
		m.fromName = valueStr
		return true
	case "to":
		m.to = m.parseAddressList(valueStr)
		return true
	case "cc":
		m.cc = m.parseAddressList(valueStr)
		return true
	case "bcc":
		m.bcc = m.parseAddressList(valueStr)
		return true
	case "subject":
		m.subject = valueStr
		return true
	case "body", "message", "textbody":
		m.body = valueStr
		m.isHTML = false
		return true
	case "htmlbody":
		m.body = valueStr
		m.isHTML = true
		return true
	case "ishtml":
		m.isHTML = value.Type == VTBool && value.Num != 0
		return true
	case "bodyformat", "mailformat":
		m.isHTML = m.vm.asInt(value) == 0
		return true
	}
	return false
}

// DispatchMethod provides O(1) string matching resolution.
func (m *G3Mail) DispatchMethod(methodName string, args []Value) Value {
	switch strings.ToLower(methodName) {
	case "addaddress", "addrecipient", "addto":
		if len(args) > 0 {
			addr := strings.TrimSpace(args[0].String())
			if addr != "" {
				m.to = append(m.to, addr)
			}
		}
		return NewBool(true)

	case "addcc":
		if len(args) > 0 {
			addr := strings.TrimSpace(args[0].String())
			if addr != "" {
				m.cc = append(m.cc, addr)
			}
		}
		return NewBool(true)

	case "addbcc":
		if len(args) > 0 {
			addr := strings.TrimSpace(args[0].String())
			if addr != "" {
				m.bcc = append(m.bcc, addr)
			}
		}
		return NewBool(true)

	case "send":
		if len(args) >= 3 {
			// CDONTS.NewMail style args: To, Subject, Body
			m.to = m.parseAddressList(args[0].String())
			m.subject = args[1].String()
			m.body = args[2].String()
			m.isHTML = false
		}
		return m.sendInternal()

	case "clear":
		m.to = []string{}
		m.cc = []string{}
		m.bcc = []string{}
		m.subject = ""
		m.body = ""
		m.isHTML = false
		return NewBool(true)
	}

	return NewEmpty()
}

func (m *G3Mail) sendInternal() Value {
	host := strings.TrimSpace(m.host)
	port := m.port
	username := strings.TrimSpace(m.username)
	password := m.password
	from := strings.TrimSpace(m.from)

	if host == "" || port <= 0 || username == "" || password == "" || from == "" {
		envHost, envPort, envUser, envPass, envFrom, envErr := m.getSMTPConfigFromEnv()
		if envErr != nil {
			return NewString(fmt.Sprintf("Error: %v", envErr))
		}
		if host == "" {
			host = envHost
		}
		if port <= 0 {
			port = envPort
		}
		if username == "" {
			username = envUser
		}
		if password == "" {
			password = envPass
		}
		if from == "" {
			from = envFrom
		}
	}

	allRecipients := append([]string{}, m.to...)
	allRecipients = append(allRecipients, m.cc...)
	allRecipients = append(allRecipients, m.bcc...)
	if len(allRecipients) == 0 {
		return NewString("Error: Missing recipients")
	}

	to := strings.Join(allRecipients, ",")

	msg := gomail.NewMessage()
	msg.SetHeader("From", from)
	msg.SetHeader("To", to)
	msg.SetHeader("Subject", m.subject)

	if m.isHTML {
		msg.SetBody("text/html", m.body)
	} else {
		msg.SetBody("text/plain", m.body)
	}

	d := gomail.NewDialer(host, port, username, password)

	if err := d.DialAndSend(msg); err != nil {
		return NewString(fmt.Sprintf("Error sending email: %v", err))
	}
	return NewBool(true)
}

func (m *G3Mail) parseAddressList(value string) []string {
	trimmed := strings.TrimSpace(value)
	if trimmed == "" {
		return []string{}
	}
	parts := strings.FieldsFunc(trimmed, func(r rune) bool {
		return r == ';' || r == ','
	})
	out := make([]string, 0, len(parts))
	for _, part := range parts {
		candidate := strings.TrimSpace(part)
		if candidate != "" {
			out = append(out, candidate)
		}
	}
	return out
}

func (m *G3Mail) getSMTPConfigFromEnv() (host string, port int, username, password, from string, err error) {
	host = strings.TrimSpace(os.Getenv("SMTP_HOST"))
	portStr := strings.TrimSpace(os.Getenv("SMTP_PORT"))
	username = strings.TrimSpace(os.Getenv("SMTP_USER"))
	password = strings.TrimSpace(os.Getenv("SMTP_PASS"))
	from = strings.TrimSpace(os.Getenv("SMTP_FROM"))

	if host == "" || portStr == "" || username == "" || password == "" || from == "" {
		return "", 0, "", "", "", fmt.Errorf("SMTP environment variables not set")
	}

	parsedPort, parseErr := strconv.Atoi(portStr)
	if parseErr != nil {
		return "", 0, "", "", "", fmt.Errorf("SMTP_PORT is not a valid number")
	}

	return host, parsedPort, username, password, from, nil
}
