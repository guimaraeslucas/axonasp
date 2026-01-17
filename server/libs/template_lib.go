package asp

import (
	"bytes"
	"fmt"
	"html/template"
	"path/filepath"
	"strings"
)

// G3TEMPLATE implements Component interface for Template operations
type G3TEMPLATE struct {
	ctx *ExecutionContext
}

func (t *G3TEMPLATE) GetProperty(name string) interface{} {
	return nil
}

func (t *G3TEMPLATE) SetProperty(name string, value interface{}) {}

func (t *G3TEMPLATE) CallMethod(name string, args []interface{}) interface{} {
	if len(args) < 1 {
		return "Error: Template path required"
	}

	getStr := func(i int) string {
		if i >= len(args) {
			return ""
		}
		return fmt.Sprintf("%v", args[i])
	}

	method := strings.ToLower(name)

	switch method {
	case "render":
		relPath := getStr(0)
		fullPath := t.ctx.Server_MapPath(relPath)

		rootDir, _ := filepath.Abs(t.ctx.RootDir)
		absPath, _ := filepath.Abs(fullPath)

		if !strings.HasPrefix(strings.ToLower(absPath), strings.ToLower(rootDir)) {
			return "Error: Access Denied"
		}

		var data interface{}
		if len(args) > 1 {
			data = args[1]
		}

		tmpl, err := template.ParseFiles(fullPath)
		if err != nil {
			return fmt.Sprintf("Error parsing template: %v", err)
		}

		var buf bytes.Buffer
		if err := tmpl.Execute(&buf, data); err != nil {
			return fmt.Sprintf("Error executing template: %v", err)
		}

		return buf.String()
	}
	return nil
}

// TemplateHelper handles Go template rendering (Legacy)
func TemplateHelper(method string, args []string, ctx *ExecutionContext) interface{} {
	lib := &G3TEMPLATE{ctx: ctx}

	var ifaceArgs []interface{}
	for _, a := range args {
		ifaceArgs = append(ifaceArgs, EvaluateExpression(a, ctx))
	}

	return lib.CallMethod(method, ifaceArgs)
}
