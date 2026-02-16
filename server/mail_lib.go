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

import (
	"fmt"
	"os"
	"strconv"
	"strings"

	"gopkg.in/gomail.v2"
)

// G3MAIL implements Component interface for Mail operations
type G3MAIL struct {
	ctx *ExecutionContext
}

func (m *G3MAIL) GetProperty(name string) interface{} {
	return nil
}

func (m *G3MAIL) SetProperty(name string, value interface{}) {}

func (m *G3MAIL) CallMethod(name string, args ...interface{}) interface{} {
	if len(args) < 1 {
		return nil
	}

	getStr := func(i int) string {
		if i >= len(args) {
			return ""
		}
		return fmt.Sprintf("%v", args[i])
	}

	getInt := func(i int) int {
		if i >= len(args) {
			return 0
		}
		val := args[i]
		if iVal, ok := val.(int); ok {
			return iVal
		}
		if sVal, ok := val.(string); ok {
			if v, err := strconv.Atoi(sVal); err == nil {
				return v
			}
		}
		return 0
	}

	getBool := func(i int) bool {
		if i >= len(args) {
			return false
		}
		val := args[i]
		if b, ok := val.(bool); ok {
			return b
		}
		s := strings.ToLower(fmt.Sprintf("%v", val))
		return s == "true" || s == "1"
	}

	method := strings.ToLower(name)

	switch method {
	case "send":
		if len(args) < 8 {
			return "Error: Insufficient arguments"
		}
		host := getStr(0)
		port := getInt(1)
		username := getStr(2)
		password := getStr(3)
		from := getStr(4)
		to := getStr(5)
		subject := getStr(6)
		body := getStr(7)
		isHtml := false
		if len(args) > 8 {
			isHtml = getBool(8)
		}

		return sendMailInternal(host, port, username, password, from, to, subject, body, isHtml)

	case "sendstandard", "sendfromenv":
		if len(args) < 3 {
			return "Error: Insufficient arguments"
		}

		host := os.Getenv("SMTP_HOST")
		portStr := os.Getenv("SMTP_PORT")
		username := os.Getenv("SMTP_USER")
		password := os.Getenv("SMTP_PASS")
		from := os.Getenv("SMTP_FROM")

		if host == "" || portStr == "" || username == "" || password == "" || from == "" {
			return "Error: SMTP environment variables not set"
		}

		port, err := strconv.Atoi(portStr)
		if err != nil {
			return "Error: SMTP_PORT is not a valid number"
		}

		to := getStr(0)
		subject := getStr(1)
		body := getStr(2)
		isHtml := false
		if len(args) > 3 {
			isHtml = getBool(3)
		}

		return sendMailInternal(host, port, username, password, from, to, subject, body, isHtml)
	}

	return nil
}

// MailHelper handles email operations (Legacy)
func MailHelper(method string, args []string, ctx *ExecutionContext) interface{} {
	lib := &G3MAIL{ctx: ctx}

	var ifaceArgs []interface{}
	for _, a := range args {
		ifaceArgs = append(ifaceArgs, EvaluateExpression(a, ctx))
	}

	return lib.CallMethod(method, ifaceArgs)
}

func sendMailInternal(host string, port int, username, password, from, to, subject, body string, isHtml bool) interface{} {
	m := gomail.NewMessage()
	m.SetHeader("From", from)
	m.SetHeader("To", to)
	m.SetHeader("Subject", subject)

	if isHtml {
		m.SetBody("text/html", body)
	} else {
		m.SetBody("text/plain", body)
	}

	d := gomail.NewDialer(host, port, username, password)

	if err := d.DialAndSend(m); err != nil {
		return fmt.Sprintf("Error sending email: %v", err)
	}
	return true
}

func getSMTPConfigFromEnv() (host string, port int, username, password, from string, err error) {
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

type LegacyMailMessage struct {
	ctx      *ExecutionContext
	objName  string
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

func newLegacyMailMessage(ctx *ExecutionContext, objName string) *LegacyMailMessage {
	return &LegacyMailMessage{
		ctx:     ctx,
		objName: objName,
		to:      make([]string, 0),
		cc:      make([]string, 0),
		bcc:     make([]string, 0),
	}
}

func NewPersitsMailSender(ctx *ExecutionContext) *LegacyMailMessage {
	return newLegacyMailMessage(ctx, "Persits.MailSender")
}

func NewCDOMessage(ctx *ExecutionContext) *LegacyMailMessage {
	return newLegacyMailMessage(ctx, "CDO.Message")
}

func NewCDONTSNewMail(ctx *ExecutionContext) *LegacyMailMessage {
	return newLegacyMailMessage(ctx, "CDONTS.NewMail")
}

func (m *LegacyMailMessage) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "host", "mailhost", "smtpserver":
		return m.host
	case "port", "smtpserverport":
		return m.port
	case "username", "user", "authusername":
		return m.username
	case "password", "pass", "authpassword":
		return m.password
	case "from", "fromaddress":
		return m.from
	case "fromname":
		return m.fromName
	case "to":
		return strings.Join(m.to, ",")
	case "cc":
		return strings.Join(m.cc, ",")
	case "bcc":
		return strings.Join(m.bcc, ",")
	case "subject":
		return m.subject
	case "body", "textbody", "htmlbody", "message":
		return m.body
	case "ishtml", "bodyformat", "mailformat":
		if strings.ToLower(name) == "bodyformat" || strings.ToLower(name) == "mailformat" {
			if m.isHTML {
				return 0
			}
			return 1
		}
		return m.isHTML
	}
	return nil
}

func (m *LegacyMailMessage) SetProperty(name string, value interface{}) {
	valueStr := fmt.Sprintf("%v", value)
	switch strings.ToLower(name) {
	case "host", "mailhost", "smtpserver":
		m.host = strings.TrimSpace(valueStr)
	case "port", "smtpserverport":
		m.port = toInt(value)
	case "username", "user", "authusername":
		m.username = strings.TrimSpace(valueStr)
	case "password", "pass", "authpassword":
		m.password = valueStr
	case "from", "fromaddress":
		m.from = strings.TrimSpace(valueStr)
	case "fromname":
		m.fromName = valueStr
	case "to":
		m.to = parseAddressList(valueStr)
	case "cc":
		m.cc = parseAddressList(valueStr)
	case "bcc":
		m.bcc = parseAddressList(valueStr)
	case "subject":
		m.subject = valueStr
	case "body", "message", "textbody":
		m.body = valueStr
		m.isHTML = false
	case "htmlbody":
		m.body = valueStr
		m.isHTML = true
	case "ishtml":
		m.isHTML = toBool(value)
	case "bodyformat", "mailformat":
		m.isHTML = toInt(value) == 0
	}
}

func (m *LegacyMailMessage) CallMethod(name string, args ...interface{}) interface{} {
	switch strings.ToLower(name) {
	case "addaddress", "addrecipient", "addto":
		if len(args) > 0 {
			addr := strings.TrimSpace(fmt.Sprintf("%v", args[0]))
			if addr != "" {
				m.to = append(m.to, addr)
			}
		}
		return true
	case "addcc":
		if len(args) > 0 {
			addr := strings.TrimSpace(fmt.Sprintf("%v", args[0]))
			if addr != "" {
				m.cc = append(m.cc, addr)
			}
		}
		return true
	case "addbcc":
		if len(args) > 0 {
			addr := strings.TrimSpace(fmt.Sprintf("%v", args[0]))
			if addr != "" {
				m.bcc = append(m.bcc, addr)
			}
		}
		return true
	case "send":
		return m.send(args...)
	case "clear":
		m.to = []string{}
		m.cc = []string{}
		m.bcc = []string{}
		m.subject = ""
		m.body = ""
		m.isHTML = false
		return true
	}
	return nil
}

func (m *LegacyMailMessage) send(args ...interface{}) interface{} {
	if strings.EqualFold(m.objName, "CDONTS.NewMail") && len(args) >= 3 {
		m.to = parseAddressList(fmt.Sprintf("%v", args[0]))
		m.subject = fmt.Sprintf("%v", args[1])
		m.body = fmt.Sprintf("%v", args[2])
		m.isHTML = false
	}

	host := strings.TrimSpace(m.host)
	port := m.port
	username := strings.TrimSpace(m.username)
	password := m.password
	from := strings.TrimSpace(m.from)

	if host == "" || port <= 0 || username == "" || password == "" || from == "" {
		envHost, envPort, envUser, envPass, envFrom, envErr := getSMTPConfigFromEnv()
		if envErr != nil {
			return fmt.Sprintf("Error: %v", envErr)
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
		return "Error: Missing recipients"
	}

	to := strings.Join(allRecipients, ",")
	return sendMailInternal(host, port, username, password, from, to, m.subject, m.body, m.isHTML)
}

func parseAddressList(value string) []string {
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
