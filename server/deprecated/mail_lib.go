package asp

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

func (m *G3MAIL) CallMethod(name string, args []interface{}) interface{} {
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
