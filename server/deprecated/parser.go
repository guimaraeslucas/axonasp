package asp

import (
	"fmt"
	"math"
	"math/rand"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"
)

// ProcessIncludes recursively handles <!--#include file="..."--> and virtual
func ProcessIncludes(content string, currentDir string, rootDir string) (string, error) {
	// Regex for includes: <!-- #include file|virtual = "path" -->
	// Handle spaces loosely
	re := regexp.MustCompile(`(?i)<!--\s*#include\s+(file|virtual)\s*=\s*"([^"]+)"\s*-->`)

	var finalErr error

	result := re.ReplaceAllStringFunc(content, func(match string) string {
		submatch := re.FindStringSubmatch(match)
		if len(submatch) < 3 {
			return match
		}

		incType := strings.ToLower(submatch[1])
		incPath := submatch[2]

		var fullPath string
		if incType == "virtual" {
			// Virtual: relative to root
			// If starts with /, strip it (filepath.Join handles it but good to be clean)
			cleanPath := strings.TrimPrefix(incPath, "/")
			cleanPath = strings.TrimPrefix(cleanPath, "\\")
			// Normalize slashes
			cleanPath = filepath.FromSlash(cleanPath)
			fullPath = filepath.Join(rootDir, cleanPath)
		} else {
			// File: relative to current file directory
			cleanPath := filepath.FromSlash(incPath)
			fullPath = filepath.Join(currentDir, cleanPath)
		}

		// Read File
		incContentBytes, err := os.ReadFile(fullPath)
		if err != nil {
			finalErr = fmt.Errorf("failed to include file '%s': %v", incPath, err)
			return fmt.Sprintf("<!-- Include Error: %v -->", err)
		}

		incContent := string(incContentBytes)

		// Recursive Process
		// New currentDir is the directory of the included file
		processed, err := ProcessIncludes(incContent, filepath.Dir(fullPath), rootDir)
		if err != nil {
			finalErr = err
			return fmt.Sprintf("<!-- Recursive Include Error: %v -->", err)
		}

		return processed
	})

	return result, finalErr
}

// ResolveObjectPath traverses "Obj.Prop1.Prop2" and returns the final object/component and the last property name.
// Returns (component, lastSegment, found)
func ResolveObjectPath(ctx *ExecutionContext, fullPath string) (Component, string, bool) {
	parts := strings.Split(fullPath, ".")
	if len(parts) == 0 {
		return nil, "", false
	}

	// Resolve base object
	baseName := parts[0]
	var baseVal interface{}
	var ok bool

	if strings.EqualFold(baseName, "me") && ctx.CurrentInstance != nil {
		baseVal = ctx.CurrentInstance
		ok = true
	} else {
		baseVal, ok = ctx.Variables[strings.ToLower(baseName)]
	}

	if !ok {
		return nil, "", false
	}

	currentComp, ok := baseVal.(Component)
	if !ok {
		return nil, "", false
	}

	// Traverse intermediate properties
	for i := 1; i < len(parts)-1; i++ {
		propName := parts[i]
		val := currentComp.GetProperty(propName)
		if val == nil {
			return nil, "", false
		}

		// Next level must be a Component
		nextComp, ok := val.(Component)
		if !ok {
			return nil, "", false
		}
		currentComp = nextComp
	}

	// Return final component and the last segment
	return currentComp, parts[len(parts)-1], true
}

// ParseGlobalASA extracts code from <SCRIPT LANGUAGE=VBScript RUNAT=Server> tags
func ParseGlobalASA(content string) string {
	// Simple regex-based extraction
	// We want everything inside <script ... runat=server> ... </script>
	// Note: Case insensitive
	re := regexp.MustCompile(`(?i)<script\s+language="?vbscript"?\s+runat="?server"?\s*>([\s\S]*?)</script>`)
	matches := re.FindAllStringSubmatch(content, -1)

	var fullCode strings.Builder
	for _, m := range matches {
		if len(m) >= 2 {
			fullCode.WriteString(m[1])
			fullCode.WriteString("\n")
		}
	}
	return fullCode.String()
}

type TokenType int

const (
	TokenHTML TokenType = iota
	TokenCode
)

type Token struct {
	Type    TokenType
	Content string
	LineNum int
}

func ParseRaw(content string) []Token {
	var tokens []Token
	length := len(content)
	cursor := 0
	line := 1

	for cursor < length {
		startCode := strings.Index(content[cursor:], "<%")
		if startCode == -1 {
			tokens = append(tokens, Token{Type: TokenHTML, Content: content[cursor:], LineNum: line})
			break
		}
		if startCode > 0 {
			htmlPart := content[cursor : cursor+startCode]
			tokens = append(tokens, Token{Type: TokenHTML, Content: htmlPart, LineNum: line})
			line += strings.Count(htmlPart, "\n")
		}
		cursor += startCode + 2
		isPrint := false
		if cursor < length && content[cursor] == '=' {
			isPrint = true
			cursor++
		}
		
		// Find end code %> respecting quotes
		endCodeRel := findEndCode(content[cursor:])
		
		if endCodeRel == -1 {
			break
		}
		codePart := content[cursor : cursor+endCodeRel]
		if isPrint {
			codePart = "Response.Write(" + codePart + ")"
		}
		tokens = append(tokens, Token{Type: TokenCode, Content: codePart, LineNum: line})
		line += strings.Count(codePart, "\n")
		cursor += endCodeRel + 2
	}
	return tokens
}

// findEndCode finds the index of "%>" ignoring occurrences inside quotes
func findEndCode(s string) int {
	inQuote := false
	for i := 0; i < len(s); i++ {
		if s[i] == '"' {
			// Check for escaped quote ""
			if i+1 < len(s) && s[i+1] == '"' {
				i++ // Skip next
			} else {
				inQuote = !inQuote
			}
		}
		if !inQuote {
			if strings.HasPrefix(s[i:], "%>") {
				return i
			}
		}
	}
	return -1
}

// Helper to find operator index respecting quotes, hashes, and parentheses
func findOpIndex(s string, op string) int {
	inquote := false
	inHash := false
	depth := 0

	for i := 0; i < len(s); i++ {
		if s[i] == '"' && !inHash {
			inquote = !inquote
		}
		if s[i] == '#' && !inquote {
			inHash = !inHash
		}
		if !inquote && !inHash {
			if s[i] == '(' {
				depth++
			} else if s[i] == ')' {
				depth--
			} else if depth == 0 {
				// Check for op match
				if strings.HasPrefix(s[i:], op) {
					return i
				}
			}
		}
	}
	return -1
}

// findLastWordOpIndex finds the last top-level occurrence of a word operator
// (case-insensitive) respecting quotes, hashes, parentheses and word boundaries.
func findLastWordOpIndex(s string, ops []string) (int, string) {
	inquote := false
	inHash := false
	depth := 0
	lastIdx := -1
	matchedOp := ""
	sLower := strings.ToLower(s)

	for i := 0; i < len(s); i++ {
		if s[i] == '"' && !inHash {
			inquote = !inquote
		}
		if s[i] == '#' && !inquote {
			inHash = !inHash
		}
		if !inquote && !inHash {
			if s[i] == '(' {
				depth++
			} else if s[i] == ')' {
				depth--
			} else if depth == 0 {
				for _, op := range ops {
					opLower := strings.ToLower(op)
					if i+len(opLower) <= len(sLower) && sLower[i:i+len(opLower)] == opLower {
						// Check word boundaries: before and after must not be letter, digit or '_'
						beforeOK := (i == 0) || !((sLower[i-1] >= 'a' && sLower[i-1] <= 'z') || (sLower[i-1] >= '0' && sLower[i-1] <= '9') || sLower[i-1] == '_')
						afterIdx := i + len(opLower)
						afterOK := (afterIdx >= len(sLower)) || !((sLower[afterIdx] >= 'a' && sLower[afterIdx] <= 'z') || (sLower[afterIdx] >= '0' && sLower[afterIdx] <= '9') || sLower[afterIdx] == '_')
						if beforeOK && afterOK {
							lastIdx = i
							matchedOp = s[i : i+len(opLower)]
							i += len(opLower) - 1
							break
						}
					}
				}
			}
		}
	}
	return lastIdx, matchedOp
}

// Find LAST occurrence for left-associativity (e.g. 10 - 5 - 2)
// Also used for finding boolean ops in EvaluateCondition
func findLastOpIndex(s string, ops []string) (int, string) {
	inquote := false
	inHash := false
	depth := 0
	lastIdx := -1
	matchedOp := ""

	for i := 0; i < len(s); {
		if s[i] == '"' && !inHash {
			inquote = !inquote
		}
		if s[i] == '#' && !inquote {
			inHash = !inHash
		}
		if !inquote && !inHash {
			if s[i] == '(' {
				depth++
			} else if s[i] == ')' {
				depth--
			} else if depth == 0 {
				// Find longest matching op at this position
				bestOp := ""
				for _, op := range ops {
					if strings.HasPrefix(s[i:], op) {
						if len(op) > len(bestOp) {
							bestOp = op
						}
					}
				}

				if bestOp != "" {
					lastIdx = i
					matchedOp = bestOp
					i += len(bestOp) // Skip the operator to avoid matching inside it (e.g. > inside <>)
					continue
				}
			}
		}
		i++
	}
	return lastIdx, matchedOp
}

func splitArgs(s string) []string {
	var args []string
	var current strings.Builder
	inquote := false
	inHash := false
	depth := 0

	for i := 0; i < len(s); i++ {
		c := s[i]
		if c == '"' && !inHash {
			inquote = !inquote
		}
		if c == '#' && !inquote {
			inHash = !inHash
		}
		if !inquote && !inHash {
			if c == '(' {
				depth++
			}
			if c == ')' {
				depth--
			}
		}

		if c == ',' && !inquote && !inHash && depth == 0 {
			args = append(args, strings.TrimSpace(current.String()))
			current.Reset()
		} else {
			current.WriteByte(c)
		}
	}
	if current.Len() > 0 {
		args = append(args, strings.TrimSpace(current.String()))
	} else if len(args) > 0 {
		// Trailing comma? Or empty? "a," -> "a", ""
		// But loop finishes. If "Call()" s is "", args empty.
		// If "Call(a,)" s is "a,", args=["a"], current="" -> append ""
		if len(s) > 0 && s[len(s)-1] == ',' {
			args = append(args, "")
		}
	}
	return args
}

// splitWithSegments breaks a dotted path into segments while respecting quotes and parentheses.
func splitWithSegments(path string) []string {
	var segments []string
	var current strings.Builder
	depth := 0
	inQuote := false

	for i := 0; i < len(path); i++ {
		c := path[i]
		if c == '"' {
			inQuote = !inQuote
		}
		if !inQuote {
			if c == '(' {
				depth++
			} else if c == ')' {
				if depth > 0 {
					depth--
				}
			} else if c == '.' && depth == 0 {
				segments = append(segments, strings.TrimSpace(current.String()))
				current.Reset()
				continue
			}
		}
		current.WriteByte(c)
	}

	if current.Len() > 0 {
		segments = append(segments, strings.TrimSpace(current.String()))
	}

	return segments
}

// evaluateWithSegment resolves a single segment against a base object.
func evaluateWithSegment(current interface{}, segment string, ctx *ExecutionContext) interface{} {
	if current == nil {
		return nil
	}

	segment = strings.TrimSpace(segment)
	if segment == "" {
		return current
	}

	if idx := strings.Index(segment, "("); idx > -1 && strings.HasSuffix(segment, ")") {
		name := strings.TrimSpace(segment[:idx])
		argStr := segment[idx+1 : len(segment)-1]
		argParts := splitArgs(argStr)
		args := make([]interface{}, 0, len(argParts))
		var binders []ByRefSetter
		if len(argParts) > 0 {
			binders = make([]ByRefSetter, len(argParts))
		}
		for i, a := range argParts {
			args = append(args, EvaluateExpression(a, ctx))
			if setter, ok := ResolveByRefSetter(ctx, a); ok {
				binders[i] = setter
			}
		}

		switch cur := current.(type) {
		case *ClassInstance:
			if ctx != nil && ctx.Engine != nil {
				return ctx.Engine.ExecuteClassMethod(cur, name, PropGet, args, binders)
			}
			return nil
		case Component:
			return cur.CallMethod(name, args)
		case map[string]interface{}:
			if strings.EqualFold(name, "item") && len(args) > 0 {
				key := fmt.Sprintf("%v", args[0])
				if val, ok := cur[key]; ok {
					return val
				}
				if val, ok := cur[strings.ToLower(key)]; ok {
					return val
				}
				return nil
			}
		case []interface{}:
			if name == "" && len(args) > 0 {
				if idxVal, ok := toInt(args[0]); ok && idxVal >= 0 && idxVal < len(cur) {
					return cur[idxVal]
				}
			}
		}

		return nil
	}

	switch cur := current.(type) {
	case *ClassInstance:
		lower := strings.ToLower(segment)
		if val, ok := cur.Variables[lower]; ok {
			return val
		}
		if cur.ClassDef != nil && ctx != nil && ctx.Engine != nil {
			if props, ok := cur.ClassDef.Properties[lower]; ok {
				for _, p := range props {
					if p.Type == PropGet && len(p.Params) == 0 {
						return ctx.Engine.ExecuteClassMethod(ctx.CurrentInstance, p.Name, PropGet, []interface{}{}, nil)
					}
				}
			}
			if _, ok := cur.ClassDef.Methods[lower]; ok {
				return ctx.Engine.ExecuteClassMethod(ctx.CurrentInstance, lower, PropGet, []interface{}{}, nil)
			}
		}
		return nil
	case Component:
		if val := cur.GetProperty(segment); val != nil {
			return val
		}
		return cur.CallMethod(segment, []interface{}{})
	case map[string]interface{}:
		if val, ok := cur[segment]; ok {
			return val
		}
		if val, ok := cur[strings.ToLower(segment)]; ok {
			return val
		}
	}

	return nil
}

func evaluateWithPath(base interface{}, path string, ctx *ExecutionContext) interface{} {
	segments := splitWithSegments(path)
	current := base
	for _, seg := range segments {
		if seg == "" {
			continue
		}
		current = evaluateWithSegment(current, seg, ctx)
		if current == nil {
			return nil
		}
	}
	return current
}

func setWithSegment(current interface{}, segment string, value interface{}, ctx *ExecutionContext) bool {
	if current == nil {
		return false
	}

	segment = strings.TrimSpace(segment)
	if segment == "" {
		return false
	}

	if idx := strings.Index(segment, "("); idx > -1 && strings.HasSuffix(segment, ")") {
		name := strings.TrimSpace(segment[:idx])
		argStr := segment[idx+1 : len(segment)-1]
		argParts := splitArgs(argStr)
		args := make([]interface{}, 0, len(argParts))
		var binders []ByRefSetter
		if len(argParts) > 0 {
			binders = make([]ByRefSetter, len(argParts))
		}
		for i, a := range argParts {
			args = append(args, EvaluateExpression(a, ctx))
			if setter, ok := ResolveByRefSetter(ctx, a); ok {
				binders[i] = setter
			}
		}

		switch cur := current.(type) {
		case map[string]interface{}:
			if strings.EqualFold(name, "item") && len(args) > 0 {
				key := fmt.Sprintf("%v", args[0])
				cur[key] = value
				return true
			}
		case []interface{}:
			if name == "" && len(args) > 0 {
				if idxVal, ok := toInt(args[0]); ok && idxVal >= 0 && idxVal < len(cur) {
					cur[idxVal] = value
					return true
				}
			}
		case *ClassInstance:
			if ctx != nil && ctx.Engine != nil {
				ctx.Engine.ExecuteClassMethod(cur, name, PropLet, append(args, value), binders)
				return true
			}
		}

		return false
	}

	switch cur := current.(type) {
	case *ClassInstance:
		lower := strings.ToLower(segment)
		if cur.ClassDef != nil {
			if _, ok := cur.ClassDef.Variables[lower]; ok {
				cur.Variables[lower] = value
				return true
			}
			if ctx != nil && ctx.Engine != nil {
				if props, ok := cur.ClassDef.Properties[lower]; ok {
					for _, p := range props {
						if p.Type == PropLet || p.Type == PropSet {
							ctx.Engine.ExecuteClassMethod(cur, p.Name, p.Type, []interface{}{value}, nil)
							return true
						}
					}
				}
			}
		}
	case Component:
		cur.SetProperty(segment, value)
		return true
	case map[string]interface{}:
		cur[segment] = value
		return true
	}

	if tryCallSetProperty(current, segment, value) {
		return true
	}

	return false
}

func setWithPath(base interface{}, path string, value interface{}, ctx *ExecutionContext) bool {
	segments := splitWithSegments(path)
	if len(segments) == 0 {
		return false
	}

	current := base
	for i := 0; i < len(segments)-1; i++ {
		seg := segments[i]
		if seg == "" {
			continue
		}
		current = evaluateWithSegment(current, seg, ctx)
		if current == nil {
			return false
		}
	}

	return setWithSegment(current, segments[len(segments)-1], value, ctx)
}

// executePrivateFunction executes a user-defined private function from Engine.Labels
// It handles parameter binding, byref setters, and return value evaluation
func executePrivateFunction(ctx *ExecutionContext, funcNameLower string, args []string, proc Procedure) interface{} {
	if ctx.Engine == nil {
		return nil
	}

	// Mark this function as executing to prevent recursion
	if ctx.ExecutingFunctions == nil {
		ctx.ExecutingFunctions = make(map[string]bool)
	}
	ctx.ExecutingFunctions[funcNameLower] = true
	defer func() {
		delete(ctx.ExecutingFunctions, funcNameLower)
	}()

	// Save current execution context
	savedVars := make(map[string]interface{})
	for k, v := range ctx.Variables {
		savedVars[k] = v
	}

	// Evaluate arguments in the caller's context BEFORE switching to function context
	evaluatedArgs := make([]interface{}, len(args))
	byrefSetters := make([]ByRefSetter, len(args))
	for i, argExpr := range args {
		evaluatedArgs[i] = EvaluateExpression(argExpr, ctx)
		// Try to get a byref setter for the argument
		if setter, ok := ResolveByRefSetter(ctx, argExpr); ok {
			byrefSetters[i] = setter
		}
	}

	// Create new variable scope for the function
	ctx.Variables = make(map[string]interface{})

	// Pre-initialize the return variable (function name) so it can be assigned without Dim
	// This allows patterns like: Function pre(value)
	//                               pre = Replace(value, ...)
	//                            End Function
	// Without needing: Dim pre
	// This is standard VBScript behavior where the function name is implicitly a local variable
	ctx.Variables[funcNameLower] = ""

	// Bind parameters to their evaluated values
	for i, paramName := range proc.Params {
		if i < len(evaluatedArgs) {
			ctx.Variables[strings.ToLower(paramName)] = evaluatedArgs[i]
		}
	}

	// Save engine state
	savedPC := ctx.Engine.PC
	savedLine := ctx.Engine.CurrentLine

	// Set execution to start of function (SKIP the Function definition line itself)
	ctx.Engine.PC = proc.LineNum + 1
	ctx.Engine.CurrentLine = proc.LineNum

	// Save call stack and create new one for this function
	oldCallStack := ctx.Engine.CallStack
	ctx.Engine.CallStack = []CallFrame{
		{
			ReturnPC:   savedPC,
			ParamNames: proc.Params,
			Setters:    byrefSetters,
		},
	}

	// Mark single procedure mode
	oldSingleProcMode := ctx.Engine.SingleProcMode
	ctx.Engine.SingleProcMode = true

	// Execute: The Engine.Run() will continue until CallStack is empty or a RETURN is hit
	ctx.Engine.Run(ctx)

	// Get the return value from the function name variable
	returnVal := ctx.Variables[strings.ToLower(funcNameLower)]

	// Capture modified parameter values BEFORE restoring context
	modifiedParams := make([]interface{}, len(proc.Params))
	for i, paramName := range proc.Params {
		modifiedParams[i] = ctx.Variables[strings.ToLower(paramName)]
	}

	// Restore caller's context
	ctx.Variables = savedVars
	ctx.Engine.PC = savedPC
	ctx.Engine.CurrentLine = savedLine
	ctx.Engine.CallStack = oldCallStack
	ctx.Engine.SingleProcMode = oldSingleProcMode

	// Apply byref setters to push changes back to caller (AFTER context restore)
	for i, setter := range byrefSetters {
		if setter != nil && i < len(modifiedParams) {
			setter(modifiedParams[i])
		}
	}

	return returnVal
}

func EvaluateExpression(expr string, ctx *ExecutionContext) (result interface{}) {
	// Error Recovery for On Error Resume Next
	defer func() {
		if r := recover(); r != nil {
			if ctx != nil && ctx.OnErrorResumeNext {
				// Populate Err object with recovered information
				if ctx.Err == nil {
					ctx.Err = &ASPError{}
				}
				ctx.Err.Number = 500
				ctx.Err.Description = fmt.Sprintf("Runtime Error: %v", r)
				ctx.Err.Source = "Microsoft VBScript runtime error"
				ctx.Err.HelpFile = ""
				ctx.Err.HelpContext = 0
				result = nil // Return nil/empty on error
			} else {
				// Re-panic if error handling is not active
				panic(r)
			}
		}
	}()

	expr = strings.TrimSpace(expr)

	// With-context implicit object access
	if strings.HasPrefix(expr, ".") {
		if ctx != nil && ctx.Engine != nil && len(ctx.Engine.WithStack) > 0 {
			base := ctx.Engine.WithStack[len(ctx.Engine.WithStack)-1]
			if base != nil {
				return evaluateWithPath(base, strings.TrimSpace(expr[1:]), ctx)
			}
			return nil
		}
	}

	// VBScript Keywords (Check first!)
	if strings.EqualFold(expr, "True") {
		return true
	}
	if strings.EqualFold(expr, "False") {
		return false
	}
	if strings.EqualFold(expr, "Null") {
		return nil
	}
	if strings.EqualFold(expr, "Empty") {
		return ""
	}
	if strings.EqualFold(expr, "Nothing") {
		return nil
	}

	// NEW Operator (Class Instantiation)
	// Check if the expression actually starts with "new" as a distinct word
	exprLower := strings.ToLower(expr)
	if strings.HasPrefix(exprLower, "new ") || strings.HasPrefix(exprLower, "new\t") || exprLower == "new" {
		parts := strings.Fields(expr)
		if len(parts) >= 2 && strings.EqualFold(parts[0], "new") {
			className := parts[1]

			// 1. Check User-Defined Classes
			if ctx != nil && ctx.GlobalClasses != nil {
				if classDef, ok := ctx.GlobalClasses[strings.ToLower(className)]; ok {
					inst := NewClassInstance(classDef, ctx)
					// Call Class_Initialize
					if ctx.Engine != nil {
						ctx.Engine.ExecuteClassMethod(inst, "Class_Initialize", PropGet, []interface{}{}, nil)
					}
					return inst
				}
			}

			// 2. Check Built-in Classes (e.g. RegExp)
			if strings.EqualFold(className, "regexp") {
				return NewRegExpObject()
			}

			// If we are here, it IS a "New X" call but X is unknown.
			// We should return nil (Object not found).
			return nil
		}
	}

	// VBScript Constants
	if strings.EqualFold(expr, "vbCrLf") || strings.EqualFold(expr, "vbNewLine") {
		return "\r\n"
	}
	if strings.EqualFold(expr, "vbCr") {
		return "\r"
	}
	if strings.EqualFold(expr, "vbLf") {
		return "\n"
	}
	if strings.EqualFold(expr, "vbTab") {
		return "\t"
	}

	// Constants (User Defined)
	if ctx.Constants != nil {
		if val, ok := ctx.Constants[strings.ToLower(expr)]; ok {
			return val
		}
	}

	// Built-in Functions (Check second)
	if strings.EqualFold(expr, "Now") || strings.EqualFold(expr, "Now()") {
		return time.Now().Format("02/01/2006 15:04:05")
	}
	if strings.EqualFold(expr, "Date") || strings.EqualFold(expr, "Date()") {
		return time.Now().Format("01/02/2006")
	}
	if strings.EqualFold(expr, "Time") || strings.EqualFold(expr, "Time()") {
		return time.Now().Format("15:04:05")
	}
	if strings.EqualFold(expr, "Year") || strings.EqualFold(expr, "Year()") {
		return time.Now().Year()
	}
	if strings.EqualFold(expr, "Month") || strings.EqualFold(expr, "Month()") {
		return int(time.Now().Month())
	}
	if strings.EqualFold(expr, "Day") || strings.EqualFold(expr, "Day()") {
		return time.Now().Day()
	}
	if strings.EqualFold(expr, "Hour") || strings.EqualFold(expr, "Hour()") {
		return time.Now().Hour()
	}
	if strings.EqualFold(expr, "Minute") || strings.EqualFold(expr, "Minute()") {
		return time.Now().Minute()
	}
	if strings.EqualFold(expr, "Second") || strings.EqualFold(expr, "Second()") {
		return time.Now().Second()
	}
	if strings.EqualFold(expr, "Weekday") || strings.EqualFold(expr, "Weekday()") {
		return int(time.Now().Weekday()) + 1
	}
	if strings.EqualFold(expr, "Timer") || strings.EqualFold(expr, "Timer()") {
		now := time.Now()
		midnight := time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, now.Location())
		return int(now.Sub(midnight).Seconds())
	}
	if strings.EqualFold(expr, "ScriptEngine") || strings.EqualFold(expr, "ScriptEngine()") {
		return "VBScript"
	}
	if strings.EqualFold(expr, "ScriptEngineBuildVersion") || strings.EqualFold(expr, "ScriptEngineBuildVersion()") {
		return 18702
	}
	if strings.EqualFold(expr, "ScriptEngineMajorVersion") || strings.EqualFold(expr, "ScriptEngineMajorVersion()") {
		return 5
	}
	if strings.EqualFold(expr, "ScriptEngineMinorVersion") || strings.EqualFold(expr, "ScriptEngineMinorVersion()") {
		return 8
	}

	// Request collections: return iterable Collection when referenced directly
	if strings.EqualFold(expr, "Request.Form") {
		if ctx != nil && ctx.Request != nil {
			return NewCollectionFromValues(ctx.Request.Form)
		}
		return NewCollectionFromValues(nil)
	}
	if strings.EqualFold(expr, "Request.QueryString") {
		if ctx != nil && ctx.Request != nil {
			return NewCollectionFromValues(ctx.Request.URL.Query())
		}
		return NewCollectionFromValues(nil)
	}

	// Logical word operators: Eqv, Imp, Xor (case-insensitive, word-bound)
	if idx, op := findLastWordOpIndex(expr, []string{" eqv ", " imp ", " xor "}); idx != -1 {
		left := EvaluateExpression(strings.TrimSpace(expr[:idx]), ctx)
		right := EvaluateExpression(strings.TrimSpace(expr[idx+len(op):]), ctx)
		toBool := func(v interface{}) bool {
			if v == nil {
				return false
			}
			if b, ok := v.(bool); ok {
				return b
			}
			if i, ok := v.(int); ok {
				return i != 0
			}
			s := fmt.Sprintf("%v", v)
			ls := strings.ToLower(s)
			return !(ls == "" || ls == "0" || ls == "false")
		}
		lb := toBool(left)
		rb := toBool(right)
		switch strings.TrimSpace(strings.ToLower(op)) {
		case "eqv":
			// Equivalence: true if both equal
			return lb == rb
		case "imp":
			// Implication: (not A) or B
			return (!lb) || rb
		case "xor":
			return lb != rb
		}
	}

	// Helper function for boolean conversion (used in And/Or)
	toBool := func(v interface{}) bool {
		if v == nil || v == "" {
			return false
		}
		if b, ok := v.(bool); ok {
			return b
		}
		if i, ok := v.(int); ok {
			return i != 0
		}
		if f, ok := v.(float64); ok {
			return f != 0
		}
		s := fmt.Sprintf("%v", v)
		ls := strings.ToLower(strings.TrimSpace(s))
		return !(ls == "" || ls == "0" || ls == "false")
	}

	// Logical OR (word operator)
	// In VBScript, Or can be bitwise (on numbers) or logical (on booleans)
	if idx, op := findLastWordOpIndex(expr, []string{"or"}); idx != -1 {
		left := EvaluateExpression(strings.TrimSpace(expr[:idx]), ctx)
		right := EvaluateExpression(strings.TrimSpace(expr[idx+len(op):]), ctx)

		// Try bitwise first (if both are numbers)
		lInt, lOk := toInt(left)
		rInt, rOk := toInt(right)
		if lOk && rOk {
			return lInt | rInt // Bitwise OR
		}

		// Fall back to logical OR
		lBool := toBool(left)
		rBool := toBool(right)
		return lBool || rBool
	}

	// Logical AND (word operator)
	// In VBScript, And can be bitwise (on numbers) or logical (on booleans)
	if idx, op := findLastWordOpIndex(expr, []string{"and"}); idx != -1 {
		left := EvaluateExpression(strings.TrimSpace(expr[:idx]), ctx)
		right := EvaluateExpression(strings.TrimSpace(expr[idx+len(op):]), ctx)

		// Try bitwise first (if both are numbers)
		lInt, lOk := toInt(left)
		rInt, rOk := toInt(right)
		if lOk && rOk {
			return lInt & rInt // Bitwise AND
		}

		// Fall back to logical AND
		lBool := toBool(left)
		rBool := toBool(right)
		return lBool && rBool
	}

	// Concatenation & (Lowest Precedence)
	// But skip if it's a hex (&h) or octal (&o) literal prefix
	if idx := findOpIndex(expr, "&"); idx != -1 {
		// Check if this & is part of &h or &o (hex/octal literal)
		if idx+1 < len(expr) {
			nextChar := strings.ToLower(string(expr[idx+1]))
			if nextChar == "h" || nextChar == "o" {
				// This is a hex/octal literal, not concatenation
				// Skip this & and continue parsing as number
				// (will be handled by parseVBScriptInt below)
			} else {
				// It's concatenation
				left := EvaluateExpression(expr[:idx], ctx)
				right := EvaluateExpression(expr[idx+1:], ctx)
				return fmt.Sprintf("%v%v", left, right)
			}
		} else {
			// & at end of expression, treat as concatenation with empty string
			left := EvaluateExpression(expr[:idx], ctx)
			return fmt.Sprintf("%v", left)
		}
	}

	// Arithmetic: + and - (Last occurrence)
	// But skip if the expression starts with - (negative number like -10 Mod 3)
	if idx, op := findLastOpIndex(expr, []string{"+", "-"}); idx != -1 && idx > 0 {
		left := EvaluateExpression(expr[:idx], ctx)
		right := EvaluateExpression(expr[idx+1:], ctx)

		lInt, lOk := toInt(left)
		rInt, rOk := toInt(right)

		if op == "+" {
			if lOk && rOk {
				return lInt + rInt
			}
			// Treat nil as 0 for addition if one side is int
			if left == nil && rOk {
				return rInt
			}
			if right == nil && lOk {
				return lInt
			}
			return fmt.Sprintf("%v%v", left, right)
		} else {
			if lOk && rOk {
				return lInt - rInt
			}
			return 0 // Error or NaN
		}
	}

	// Exponentiation ^ (High Precedence, similar to Multiplication)
	if idx := findOpIndex(expr, "^"); idx != -1 {
		left := EvaluateExpression(expr[:idx], ctx)
		right := EvaluateExpression(expr[idx+1:], ctx)
		lFloat := getFloatVal(left)
		rFloat := getFloatVal(right)
		return math.Pow(lFloat, rFloat)
	}

	// Multiplication (*), Division (/), Integer Division (\), and Modulo (Mod)
	// Handle both symbolic and word operators
	if idx, op := findLastOpIndex(expr, []string{"*", "/", "\\"}); idx != -1 {
		left := EvaluateExpression(expr[:idx], ctx)
		right := EvaluateExpression(expr[idx+1:], ctx)
		lInt, lOk := toInt(left)
		rInt, rOk := toInt(right)

		switch op {
		case "*":
			if lOk && rOk {
				return lInt * rInt
			}
			return 0
		case "/":
			if lOk && rOk {
				if rInt == 0 {
					panic("runtime error: float divide by zero")
				}
				// Regular division returns float
				return float64(lInt) / float64(rInt)
			}
			return 0.0
		case "\\":
			if lOk && rOk {
				if rInt == 0 {
					panic("runtime error: integer divide by zero")
				}
				return lInt / rInt
			}
			return 0
		}
	}

	// Modulo (Mod) - word operator with word boundaries
	if idx, op := findLastWordOpIndex(expr, []string{"mod"}); idx != -1 {
		left := EvaluateExpression(strings.TrimSpace(expr[:idx]), ctx)
		right := EvaluateExpression(strings.TrimSpace(expr[idx+len(op):]), ctx)
		lInt, lOk := toInt(left)
		rInt, rOk := toInt(right)

		if lOk && rOk {
			if rInt == 0 {
				panic("runtime error: modulo by zero")
			}
			return lInt % rInt
		}
		return 0
	}

	// String literal
	if strings.HasPrefix(expr, "\"") && strings.HasSuffix(expr, "\"") {
		if len(expr) >= 2 {
			val := expr[1 : len(expr)-1]
			// CORREÇÃO: VBScript usa "" para representar uma única aspa "
			// Isso é fundamental para strings JSON escritas dentro do ASP
			return strings.ReplaceAll(val, "\"\"", "\"")
		}
		return ""
	}

	// Date Literal (#mm/dd/yyyy#)
	if strings.HasPrefix(expr, "#") && strings.HasSuffix(expr, "#") {
		if len(expr) >= 2 {
			val := expr[1 : len(expr)-1]
			val = strings.TrimSpace(val)
			// Try parsing common formats
			// Default US Format for literals
			formats := []string{
				"1/2/2006",
				"01/02/2006",
				"2006-01-02",
				"1/2/06",
				"01/02/06",
				"15:04:05",
				"3:04:05 PM",
				"01/02/2006 15:04:05",
				"1/2/2006 3:04:05 PM",
			}

			for _, f := range formats {
				if t, err := time.Parse(f, val); err == nil {
					// Use a dedicated type or just formatted string?
					// VBScript dates are Variants (doubles), but here we often pass strings or time.Time.
					// Let's pass time.Time if possible, but many functions expect string or check type.
					// Existing functions (Year, Month) check for time.Time?
					// EvaluateStandardFunction Year(): calls time.Now().Year().
					// Wait, Year(x) calls getStr() -> EvaluateExpression -> string.
					// So passing time.Time might break things if toInt/getStr doesn't handle it.
					// Let's check `getStr` in `parser.go`.
					// `getStr` uses `fmt.Sprintf("%v", val)`. time.Time stringifies nicely.
					// But `dateadd` etc parse string.
					// So it is safe to return time.Time.
					return t
				}
			}
			// If parse fails, treat as string? Or Error?
			// VBScript error: Type mismatch usually if not a date.
			// But let's return it as string if we can't parse, maybe user meant something else?
			// Actually #...# is exclusively date.
			return val
		}
	}

	// Number (including Hex &hFF and Octal &o77)
	if i, ok := parseVBScriptInt(expr); ok {
		return i
	}

	// Parentheses (Group) - e.g. (1+2)
	// Check if the whole string is wrapped in parens
	if strings.HasPrefix(expr, "(") && strings.HasSuffix(expr, ")") {
		// Verify balancing
		depth := 0
		matchIdx := -1
		for i, c := range expr {
			if c == '(' {
				depth++
			}
			if c == ')' {
				depth--
				if depth == 0 && matchIdx == -1 {
					matchIdx = i
				}
			}
		}
		// If the closing paren that matches the start is at the end, strip them
		if matchIdx == len(expr)-1 {
			return EvaluateExpression(expr[1:len(expr)-1], ctx)
		}
	}

	// Variable (Case Insensitive Lookup)
	lowerExpr := strings.ToLower(expr)

	// Me Keyword
	if lowerExpr == "me" && ctx.CurrentInstance != nil {
		return ctx.CurrentInstance
	}

	// Check for Object Property Access (Obj.Prop) without parentheses
	if dotIdx := strings.Index(lowerExpr, "."); dotIdx > -1 && !strings.Contains(lowerExpr, "(") {
		// Attempt to resolve deep path (e.g. rs.Fields.Count or rs.MoveNext)
		comp, lastSegment, found := ResolveObjectPath(ctx, expr)
		if found {
			// Try GetProperty first
			val := comp.GetProperty(lastSegment)
			if val != nil {
				return val
			}
			// Fallback: It might be a method call without parens (e.g. rs.MoveNext)
			// Try CallMethod with empty args
			return comp.CallMethod(lastSegment, []interface{}{})
		}
	}

	if val, ok := ctx.Variables[lowerExpr]; ok {
		return val
	}

	// Global Variable Check
	if ctx.GlobalVariables != nil {
		if val, ok := ctx.GlobalVariables[lowerExpr]; ok {
			return val
		}
	}

	// Fallback: Check CurrentInstance Members (Implicit Scope)
	if ctx.CurrentInstance != nil {
		// 1. Check Instance Variables
		if val, ok := ctx.CurrentInstance.Variables[lowerExpr]; ok {
			return val
		}
		// 2. Check Instance Properties (Get without args)
		if props, ok := ctx.CurrentInstance.ClassDef.Properties[lowerExpr]; ok {
			for _, p := range props {
				if p.Type == PropGet && len(p.Params) == 0 {
					// Guard against recursion
					if lowerExpr != strings.ToLower(ctx.CurrentMethodName) {
						if ctx.Engine != nil {
							return ctx.Engine.ExecuteClassMethod(ctx.CurrentInstance, p.Name, PropGet, []interface{}{}, nil)
						}
					}
				}
			}
		}
		// 3. Check Instance Methods (Function without args)
		if _, ok := ctx.CurrentInstance.ClassDef.Methods[lowerExpr]; ok {
			// Prevent recursion: if the expression IS the current function name, don't re-execute.
			// The initialized local variable should have been found already.
			// This is a safeguard against the lookup failing and causing a loop.
			if lowerExpr != strings.ToLower(ctx.CurrentMethodName) {
				if ctx.Engine != nil {
					return ctx.Engine.ExecuteClassMethod(ctx.CurrentInstance, lowerExpr, PropGet, []interface{}{}, nil)
				}
			}
		}
	}

	// Try Function Call or Array Access: Name(Args)
	if idxStart := strings.Index(expr, "("); idxStart > -1 && strings.HasSuffix(expr, ")") {
		arrName := strings.TrimSpace(expr[:idxStart])
		idxExpr := expr[idxStart+1 : len(expr)-1]
		args := splitArgs(idxExpr) // Parse Args

		// 1. Funções Nativas do Servidor (Injete isso antes das Standard Functions)
		// [CORREÇÃO] O File System DEVE ser verificado AQUI, logo no início
		// FILE SYSTEM API
		if strings.HasPrefix(strings.ToLower(arrName), "file.") {
			method := strings.TrimPrefix(strings.ToLower(arrName), "file.")
			// Here we call FileSystemAPI implemented in file_lib.go
			return FileSystemAPI(method, args, ctx)
		}
		//CRYPTO API
		if strings.HasPrefix(strings.ToLower(arrName), "crypto.") {
			method := strings.TrimPrefix(strings.ToLower(arrName), "crypto.")
			return CryptoHelper(method, args, ctx)
		}
		//MAIL API
		if strings.HasPrefix(strings.ToLower(arrName), "mail.") {
			method := strings.TrimPrefix(strings.ToLower(arrName), "mail.")
			return MailHelper(method, args, ctx)
		}
		//TEMPLATE API
		if strings.HasPrefix(strings.ToLower(arrName), "template.") {
			method := strings.TrimPrefix(strings.ToLower(arrName), "template.")
			return TemplateHelper(method, args, ctx)
		}
		//JSON API
		if strings.EqualFold(arrName, "JSON.Parse") && len(args) > 0 {
			lib := &JSONLibrary{}
			return lib.Parse(fmt.Sprintf("%v", EvaluateExpression(args[0], ctx)))
		}
		if strings.EqualFold(arrName, "JSON.Stringify") && len(args) > 0 {
			lib := &JSONLibrary{}
			val := EvaluateExpression(args[0], ctx)
			return lib.Stringify(val)
		}
		if strings.EqualFold(arrName, "JSON.NewObject") {
			lib := &JSONLibrary{}
			return lib.NewObject()
		}
		if strings.EqualFold(arrName, "JSON.NewArray") {
			lib := &JSONLibrary{}
			return lib.NewArray()
		}
		//ENV
		if strings.EqualFold(arrName, "Env") {
			if len(args) > 0 {
				key := fmt.Sprintf("%v", EvaluateExpression(args[0], ctx))
				val := os.Getenv(key) // Requer import "os"
				return val
			}
			return ""
		}
		//FETCH
		// FETCH(URL, [METHOD], [BODY])
		if strings.EqualFold(arrName, "Fetch") {
			// Requer implementação auxiliar (veja abaixo)
			return FetchHelper(args, ctx)
		}

		// 1.1. Check Standard Functions (Len, Left, etc...)
		if val, ok := EvaluateStandardFunction(arrName, args, ctx); ok {
			return val
		}

		// 1.2. Check Private Functions (User-defined in Engine.Labels)
		if ctx != nil {
			funcNameLower := strings.ToLower(arrName)
			if ctx.ExecutingFunctions == nil {
				ctx.ExecutingFunctions = make(map[string]bool)
			}
			if !ctx.ExecutingFunctions[funcNameLower] { // Only if not already executing
				// First try ctx.Engine
				if ctx.Engine != nil && ctx.Engine.Labels != nil {
					if proc, ok := ctx.Engine.Labels[funcNameLower]; ok {
						// Execute private function with arguments
						return executePrivateFunction(ctx, funcNameLower, args, proc)
					}
				}
				// Fallback: if ctx.Engine is nil, create a temporary one from context
				// This handles cases where ExecutionContext was created without Engine
				// Usually happens during initialization or in certain execution paths
				if ctx.Engine == nil && ctx.GlobalClasses != nil {
					// We have context but no Engine reference yet
					// In this case, we'll skip the function lookup
					// The function should have been accessed via a proper execution path
				}
			}
		}

		// 2. Objects (Request, Session, etc)
		// Check if it's Request object before trying array
		if strings.EqualFold(arrName, "Request.QueryString") {
			if ctx.Request != nil {
				// Strip quotes from idxExpr if present
				key := strings.Trim(strings.TrimSpace(idxExpr), "\"")
				val := ctx.Request.URL.Query().Get(key)
				return val
			}
			return ""
		}
		if strings.EqualFold(arrName, "Request.Form") {
			if ctx.Request != nil {
				key := strings.Trim(strings.TrimSpace(idxExpr), "\"")
				val := ctx.Request.FormValue(key)
				return val
			}
			return ""
		}
		if strings.EqualFold(arrName, "Request.Cookies") {
			if ctx.Request != nil {
				key := strings.Trim(strings.TrimSpace(idxExpr), "\"")
				c, err := ctx.Request.Cookie(key)
				if err == nil {
					return c.Value
				}
				return ""
			}
			return ""
		}
		if strings.EqualFold(arrName, "Request.ServerVariables") {
			if ctx.Request != nil {
				key := strings.ToUpper(strings.Trim(strings.TrimSpace(idxExpr), "\""))
				switch key {
				case "REQUEST_METHOD":
					return ctx.Request.Method
				case "SERVER_NAME":
					return ctx.Request.URL.Hostname() // Host might include port
				case "HTTP_HOST":
					return ctx.Request.Host
				case "REMOTE_ADDR":
					return ctx.Request.RemoteAddr
				case "QUERY_STRING":
					return ctx.Request.URL.RawQuery
				case "URL", "PATH_INFO":
					return ctx.Request.URL.Path
				case "HTTP_USER_AGENT":
					return ctx.Request.UserAgent()
				}
				// Generic Header Lookup
				if strings.HasPrefix(key, "HTTP_") {
					headerKey := strings.ReplaceAll(key[5:], "_", "-")
					return ctx.Request.Header.Get(headerKey)
				}
				return ""
			}
			return ""
		}

		// Server Methods
		if strings.EqualFold(arrName, "Server.HTMLEncode") {
			arg := args[0]
			val := EvaluateExpression(arg, ctx)
			return ctx.Server_HTMLEncode(fmt.Sprintf("%v", val))
		}
		if strings.EqualFold(arrName, "Server.URLEncode") {
			arg := args[0]
			val := EvaluateExpression(arg, ctx)
			return ctx.Server_URLEncode(fmt.Sprintf("%v", val))
		}
		if strings.EqualFold(arrName, "Server.MapPath") {
			arg := args[0]
			val := EvaluateExpression(arg, ctx)
			return ctx.Server_MapPath(fmt.Sprintf("%v", val))
		}

		// ERR METHODS
		// ERR METHODS
		if strings.EqualFold(arrName, "Err.Raise") {
			// Err.Raise(Number, Source, Description, HelpFile, HelpContext)
			if len(args) > 0 {
				num, _ := toInt(EvaluateExpression(args[0], ctx))
				src := ""
				desc := ""
				helpFile := ""
				helpCtx := 0

				if len(args) > 1 {
					src = fmt.Sprintf("%v", EvaluateExpression(args[1], ctx))
				}
				if len(args) > 2 {
					desc = fmt.Sprintf("%v", EvaluateExpression(args[2], ctx))
				}
				if len(args) > 3 {
					helpFile = fmt.Sprintf("%v", EvaluateExpression(args[3], ctx))
				}
				if len(args) > 4 {
					helpCtx, _ = toInt(EvaluateExpression(args[4], ctx))
				}

				ctx.Err.Raise(num, src, desc, helpFile, helpCtx)

				// If On Error Resume Next is NOT active, raise a runtime panic to be handled by defer in EvaluateExpression
				if ctx == nil || !ctx.OnErrorResumeNext {
					panic(fmt.Sprintf("Err.Raise: %d - %s", num, desc))
				}

			}
			return nil
		}
		if strings.EqualFold(arrName, "Err.Clear") {
			ctx.Err.Clear()
			return nil
		}

		if strings.EqualFold(arrName, "Server.CreateObject") {

			if len(args) > 0 {

				progID := fmt.Sprintf("%v", EvaluateExpression(args[0], ctx))

				return ComponentFactory(progID, ctx)

			}

			return nil

		}

		if strings.EqualFold(arrName, "Server.Execute") || strings.EqualFold(arrName, "Server.Transfer") {
			if len(args) > 0 {
				path := fmt.Sprintf("%v", EvaluateExpression(args[0], ctx))
				fullPath := ctx.Server_MapPath(path)

				contentBytes, err := os.ReadFile(fullPath)
				if err == nil {
					contentStr := string(contentBytes)
					contentStr, _ = ProcessIncludes(contentStr, filepath.Dir(fullPath), ctx.RootDir)
					tokens := ParseRaw(contentStr)
					engine := Prepare(tokens)

					// Create isolated context (Variables cleared)
					// BUT keep Session/App/Response
					subCtx := *ctx
					subCtx.Variables = make(map[string]interface{})
					// Copy debugging flag if present
					if val, ok := ctx.Variables["debug_asp_code"]; ok {
						subCtx.Variables["debug_asp_code"] = val
					}

					engine.Run(&subCtx)

					// Merge output back
					// Since subCtx.Output is a separate slice header, we must append it
					if len(subCtx.Output) > 0 {
						ctx.Output = append(ctx.Output, subCtx.Output...)
					}

					// If Transfer, end the current execution
					if strings.EqualFold(arrName, "Server.Transfer") {
						ctx.End()
					}
				}
			}
			return nil
		}
		if strings.EqualFold(arrName, "Server.ScriptTimeout") {
			return ctx.ScriptTimeout
		}
		if strings.EqualFold(arrName, "Server.GetLastError") {
			return nil // Stub return ASPError object
		}
		//Custom Server Method
		if strings.EqualFold(arrName, "Server.Name") {
			return "G3pix AxonASP" // Stub return ASPError object
		}

		// ObjectContext Methods
		if strings.EqualFold(arrName, "ObjectContext.SetComplete") {
			if ctx.ObjectContext != nil {
				ctx.ObjectContext.SetComplete()
			}
			return nil
		}
		if strings.EqualFold(arrName, "ObjectContext.SetAbort") {
			if ctx.ObjectContext != nil {
				ctx.ObjectContext.SetAbort()
			}
			return nil
		}

		// Response Methods
		if strings.EqualFold(arrName, "Response.AddHeader") {
			if len(args) >= 2 {
				name := fmt.Sprintf("%v", EvaluateExpression(args[0], ctx))
				val := fmt.Sprintf("%v", EvaluateExpression(args[1], ctx))
				ctx.AddHeader(name, val)
			}
			return nil
		}
		if strings.EqualFold(arrName, "Response.BinaryWrite") {
			if len(args) > 0 {
				val := EvaluateExpression(args[0], ctx)
				if b, ok := val.([]byte); ok {
					ctx.BinaryWrite(b)
				} else {
					ctx.Write(fmt.Sprintf("%v", val))
				}
			}
			return nil
		}
		if strings.EqualFold(arrName, "Response.AppendToLog") {
			if len(args) > 0 {
				val := fmt.Sprintf("%v", EvaluateExpression(args[0], ctx))
				ctx.AppendToLog(val)
			}
			return nil
		}
		if strings.EqualFold(arrName, "Response.Redirect") {
			if len(args) > 0 {
				val := EvaluateExpression(args[0], ctx)
				ctx.Redirect(fmt.Sprintf("%v", val))
			}
			return nil
		}
		if strings.EqualFold(arrName, "Response.End") {
			ctx.End()
			return nil
		}

		// Response Properties (Getters)
		if strings.EqualFold(arrName, "Response.Expires") {
			return ctx.ResponseState.Expires
		}
		if strings.EqualFold(arrName, "Response.ExpiresAbsolute") {
			return ctx.ResponseState.ExpiresAbsolute
		}
		if strings.EqualFold(arrName, "Response.CacheControl") {
			return ctx.ResponseState.CacheControl
		}
		if strings.EqualFold(arrName, "Response.Charset") {
			return ctx.ResponseState.Charset
		}
		if strings.EqualFold(arrName, "Response.IsClientConnected") {
			return ctx.IsClientConnected()
		}
		if strings.EqualFold(arrName, "Response.PICS") {
			return ctx.ResponseState.PICS
		}

		// Document.Write (Safe Output similar to Response.Write)
		if strings.EqualFold(arrName, "Document.Write") {
			if len(args) > 0 {
				val := EvaluateExpression(args[0], ctx)
				strVal := fmt.Sprintf("%v", val)
				safeVal := ctx.Server_HTMLEncode(strVal)
				ctx.Write(safeVal)
			}
			return nil
		}

		if strings.EqualFold(arrName, "Request.TotalBytes") {
			return ctx.Request.ContentLength // Approximation
		}
		if strings.EqualFold(arrName, "Request.BinaryRead") {
			// BinaryRead(Count) - reads Count bytes from request body
			if len(args) < 1 {
				return nil
			}
			count := 0
			if countVal := EvaluateExpression(args[0], ctx); countVal != nil {
				if c, ok := countVal.(int); ok {
					count = c
				}
			}
			if count <= 0 || ctx.Request.Body == nil {
				return []byte{}
			}
			// Read up to count bytes
			buffer := make([]byte, count)
			n, err := ctx.Request.Body.Read(buffer)
			if err != nil && err.Error() != "EOF" {
				return []byte{}
			}
			return buffer[:n]
		}

		if strings.EqualFold(arrName, "Session") {
			keyExpr := strings.TrimSpace(idxExpr)
			// CORREÇÃO: Removemos a otimização manual propensa a erros.
			// Avaliamos a expressão da chave da mesma forma que o resto do sistema.
			val := EvaluateExpression(keyExpr, ctx)
			key := fmt.Sprintf("%v", val)
			return ctx.Session.Get(key)
		}
		if strings.EqualFold(arrName, "Application") {
			keyExpr := strings.TrimSpace(idxExpr)

			// CORREÇÃO: Força a avaliação padrão da chave.
			// Se keyExpr for "itemKey", EvaluateExpression vai buscar na variável e retornar "TestA".
			// Se for "TestA" (com aspas), vai retornar TestA.
			val := EvaluateExpression(keyExpr, ctx)
			key := fmt.Sprintf("%v", val)
			return ctx.Application.Get(key)
		}

		// 3. ACESSO DINÂMICO A OBJETOS E ARRAYS (A GRANDE MUDANÇA)

		// 3.1 COMPONENT METHOD CALLS (Obj.Method(Args) or Obj.Prop.Method(Args))
		// Check if arrName contains a dot (e.g. "Dict.Add" or "rs.Fields.Item")
		if strings.Contains(arrName, ".") {
			// Resolve path to the final component and method name
			comp, methodName, found := ResolveObjectPath(ctx, arrName)
			if found {
				// Prepare args
				var compArgs []interface{}
				for _, a := range args {
					compArgs = append(compArgs, EvaluateExpression(a, ctx))
				}
				return comp.CallMethod(methodName, compArgs)
			}
			// If not found via ResolveObjectPath (maybe it's not a component in Variables?),
			// it falls through to old logic or fails.
			// Old logic below tries "Variables[varName]".
		}

		// Verifica se a variável "arrName" existe e é um Objeto ou Array Go
		if baseObj, exists := ctx.Variables[strings.ToLower(arrName)]; exists {

			// Caso 0: Component default method (Item or Class Default)
			if comp, ok := baseObj.(Component); ok {
				if len(args) > 0 {
					// Check if it's a ClassInstance with a Default Method
					if classInst, isClass := baseObj.(*ClassInstance); isClass {
						if classInst.ClassDef.DefaultMethod != "" {
							var compArgs []interface{}
							for _, a := range args {
								compArgs = append(compArgs, EvaluateExpression(a, ctx))
							}
							// Use PropGet semantics to invoke the function/property
							// CallMethod in class_def.go handles redirection to ExecuteClassMethod
							return comp.CallMethod(classInst.ClassDef.DefaultMethod, compArgs)
						}
					}

					// Default to "Item" for Dictionary or Classes without Default
					key := EvaluateExpression(args[0], ctx)
					return comp.CallMethod("Item", []interface{}{key})
				}
			}

			// Caso 1: É um Objeto/Map (JSON Object) -> obj("key")
			if mapObj, ok := baseObj.(map[string]interface{}); ok {
				if len(args) > 0 {
					key := fmt.Sprintf("%v", EvaluateExpression(args[0], ctx))
					// Retorna o valor ou nil
					return mapObj[key]
				}
			}

			// Caso 2: É um Array/Slice (JSON Array) -> arr(0)
			if sliceObj, ok := baseObj.([]interface{}); ok {
				if len(args) > 0 {
					// Support multi-dimensional indexing: arr(0,1,2)
					current := interface{}(sliceObj)
					for _, a := range args {
						idxVal := EvaluateExpression(a, ctx)
						idx, isInt := toInt(idxVal)
						if !isInt || idx < 0 {
							return nil
						}
						switch cur := current.(type) {
						case []interface{}:
							if idx >= 0 && idx < len(cur) {
								current = cur[idx]
								continue
							}
							return nil
						default:
							return nil
						}
					}
					return current
				}
			}
		}

		// Fallback para o comportamento antigo (chaves planas como "myArr(0)")
		idxVal := EvaluateExpression(idxExpr, ctx)
		lookupKey := fmt.Sprintf("%s(%v)", arrName, idxVal)
		if val, ok := ctx.Variables[strings.ToLower(lookupKey)]; ok {
			return val
		}
	}

	if strings.EqualFold(expr, "Session.SessionID") {
		return ctx.Session.ID
	}
	if strings.EqualFold(expr, "Session.Timeout") {
		return ctx.Session.Timeout
	}

	// Err Object Properties
	if strings.EqualFold(expr, "Err") {
		return ctx.Err.Number // Default property? Usually returns Object, but for print we might want number or use Err.Number explicitly
	}
	if strings.EqualFold(expr, "Err.Number") {
		return ctx.Err.Number
	}
	if strings.EqualFold(expr, "Err.Description") {
		return ctx.Err.Description
	}
	if strings.EqualFold(expr, "Err.Source") {
		return ctx.Err.Source
	}

	// Err Object Methods (Handle as function calls if they have args, or properties if not?)
	// But EvaluateExpression handles "Name(Args)" below.
	// If "Err.Raise(1)" is passed, it goes to "Try Function Call or Array Access" block.
	// So we should check "Err.Raise" there.

	if strings.EqualFold(expr, "Response.Buffer") {
		return ctx.ResponseState.Buffer
	}
	if strings.EqualFold(expr, "Application.StaticObjects") {
		return ctx.Application.StaticObjects()
	}

	return expr
}

func EvaluateCondition(condition string, ctx *ExecutionContext) bool {
	// 1. Handle "Not (...)" wrapper first
	// This is a simple recursive check to peel off "Not " prefix if it wraps the whole expression
	// BUT be careful: "Not x Is Nothing" vs "x Is Not Nothing"
	// If the string starts with "not " (case insensitive) and assumes the rest is the boolean to invert?
	// VBScript "Not" is an operator.
	// For "Not (A Is B)", the "Is" finder below needs to NOT find the "Is" inside the parens.

	// Helper to find word operator index respecting quotes and parentheses
	findWordOpIndex := func(s string, op string) int {
		inquote := false
		depth := 0
		opLen := len(op)
		sLower := strings.ToLower(s)

		for i := 0; i <= len(s)-opLen; i++ {
			if s[i] == '"' {
				inquote = !inquote
			}
			if !inquote {
				if s[i] == '(' {
					depth++
				} else if s[i] == ')' {
					depth--
				} else if depth == 0 {
					// Check for op match
					if sLower[i:i+opLen] == op {
						return i
					}
				}
			}
		}
		return -1
	}

	// Check for "Is" operator first (case-insensitive) for Nothing/Null checks
	// " Is " and " Is Not "

	// Try " Is Not " first (longer match)
	spaceIsNotIdx := findWordOpIndex(condition, " is not ")
	spaceIsIdx := -1

	var isIdx int
	var isNot bool

	if spaceIsNotIdx != -1 {
		isIdx = spaceIsNotIdx
		isNot = true
	} else {
		spaceIsIdx = findWordOpIndex(condition, " is ")
		if spaceIsIdx != -1 {
			isIdx = spaceIsIdx
			isNot = false
		} else {
			isIdx = -1
		}
	}

	if isIdx != -1 {
		leftStr := strings.TrimSpace(condition[:isIdx])
		var rightStr string

		if isNot {
			rightStr = strings.TrimSpace(condition[isIdx+8:]) // Skip " is not "
		} else {
			rightStr = strings.TrimSpace(condition[isIdx+4:]) // Skip " is "
		}

		left := EvaluateExpression(leftStr, ctx)
		right := EvaluateExpression(rightStr, ctx)

		var match bool
		// Logic for "Is"
		if left == nil && right == nil {
			match = true // Nothing Is Nothing -> True
		} else if left == nil || right == nil {
			match = false // One is Nothing, other isn't -> False
		} else {
			// Both are not Nothing.
			// In strict VBScript, 'Is' checks object identity.
			// Here we treat it as value equality for simplicity, or handle objects if needed.
			match = (fmt.Sprintf("%v", left) == fmt.Sprintf("%v", right))
		}

		if isNot {
			return !match
		} else {
			return match
		}
	}

	// Logic for "Not" operator at start (e.g. "Not (x = y)")
	// If no "Is" operator found at top level, check for "Not"
	trimCond := strings.TrimSpace(condition)
	if strings.HasPrefix(strings.ToLower(trimCond), "not ") {
		inner := strings.TrimSpace(trimCond[4:])
		// If wrapped in parens "Not ( ... )", peel them?
		// EvaluateCondition handles parens implicitly via findLastOpIndex or recursion if we implement it.
		// For now, let's just recurse.
		return !EvaluateCondition(inner, ctx)
	}

	ops := []string{"<>", ">=", "<=", "=", ">", "<"}

	// Use findLastOpIndex to split (careful with order or just find ANY op?)
	// findLastOpIndex is fine.

	// We need to check all ops.
	// But `findLastOpIndex` takes a list.
	// However, we need to know WHICH one matched.

	idx, op := findLastOpIndex(condition, ops)

	if idx != -1 {
		leftStr := condition[:idx]
		rightStr := condition[idx+len(op):]

		left := EvaluateExpression(leftStr, ctx)
		right := EvaluateExpression(rightStr, ctx)

		toString := func(v interface{}) string { return fmt.Sprintf("%v", v) }

		switch op {
		case "=":
			return toString(left) == toString(right)
		case "<>":
			return toString(left) != toString(right)
		case ">":
			l, _ := toInt(left)
			r, _ := toInt(right)
			return l > r
		case "<":
			l, _ := toInt(left)
			r, _ := toInt(right)
			return l < r
		case ">=":
			l, _ := toInt(left)
			r, _ := toInt(right)
			return l >= r
		case "<=":
			l, _ := toInt(left)
			r, _ := toInt(right)
			return l <= r
		}
	}

	// NEW: Check for wrapping parentheses if no operator found (e.g. "(A Is B)")
	if strings.HasPrefix(condition, "(") && strings.HasSuffix(condition, ")") {
		// Verify balancing to ensure it's a wrapper, not (A) And (B)
		depth := 0
		wrapped := true
		for i := 0; i < len(condition)-1; i++ {
			switch condition[i] {
			case '(':
				depth++
			case ')':
				depth--
			}
			if depth == 0 {
				wrapped = false
				break
			}
		}
		if wrapped {
			return EvaluateCondition(condition[1:len(condition)-1], ctx)
		}
	}

	// No operator found, evaluate as boolean
	val := EvaluateExpression(condition, ctx)
	if b, ok := val.(bool); ok {
		return b
	}
	return val != nil && val != "" && val != 0
}

func toInt(v interface{}) (int, bool) {
	if v == nil {
		return 0, false
	}
	if i, ok := v.(int); ok {
		return i, true
	}
	if s, ok := v.(string); ok {
		// Use parseVBScriptInt to support hex (&h) and octal (&o) literals
		if i, ok := parseVBScriptInt(s); ok {
			return i, true
		}
	}
	return 0, false
}

func EvaluateStandardFunction(name string, args []string, ctx *ExecutionContext) (interface{}, bool) {
	name = strings.ToLower(name)

	// Helper to get int arg
	getInt := func(i int) int {
		if i >= len(args) {
			return 0
		}
		val := EvaluateExpression(args[i], ctx)
		iv, _ := toInt(val)
		return iv
	}
	// Helper to get string arg
	getStr := func(i int) string {
		if i >= len(args) {
			return ""
		}
		val := EvaluateExpression(args[i], ctx)
		return fmt.Sprintf("%v", val)
	}
	// Helper to get float arg
	getFloat := func(i int) float64 {
		if i >= len(args) {
			return 0.0
		}
		val := EvaluateExpression(args[i], ctx)
		if f, ok := val.(float64); ok {
			return f
		}
		if i, ok := val.(int); ok {
			return float64(i)
		}
		if s, ok := val.(string); ok {
			if f, err := strconv.ParseFloat(s, 64); err == nil {
				return f
			}
		}
		return 0.0
	}

	switch name {
	// Dynamic Execution Functions
	case "execute":
		// Execute(code) - executes code in local scope
		if len(args) < 1 {
			return nil, true
		}
		code := getStr(0)
		if code == "" {
			return nil, true
		}
		// Parse and execute the code
		tokens := ParseRaw(code)
		engine := Prepare(tokens)
		// Execute in current context (local scope)
		engine.Run(ctx)
		return nil, true

	case "executeglobal":
		// ExecuteGlobal(code) - executes code in global scope
		if len(args) < 1 {
			return nil, true
		}
		code := getStr(0)
		if code == "" {
			return nil, true
		}
		// Parse and execute the code
		tokens := ParseRaw(code)
		engine := Prepare(tokens)
		// Save current variables
		oldVars := ctx.Variables
		// Execute in global scope
		ctx.Variables = ctx.GlobalVariables
		engine.Run(ctx)
		// Restore variables (changes persist in GlobalVariables)
		ctx.Variables = oldVars
		return nil, true

	// String Functions
	case "len":
		if len(args) < 1 {
			return 0, true
		}
		return len(getStr(0)), true
	case "left":
		if len(args) < 2 {
			return "", true
		}
		s := getStr(0)
		n := getInt(1)
		if n > len(s) {
			n = len(s)
		}
		if n < 0 {
			n = 0
		}
		return s[:n], true
	case "right":
		if len(args) < 2 {
			return "", true
		}
		s := getStr(0)
		n := getInt(1)
		if n > len(s) {
			n = len(s)
		}
		if n < 0 {
			n = 0
		}
		return s[len(s)-n:], true
	case "mid":
		if len(args) < 2 {
			return "", true
		}
		s := getStr(0)
		start := getInt(1) - 1 // 1-based
		length := len(s)
		if len(args) >= 3 {
			length = getInt(2)
		}
		if start < 0 {
			start = 0
		} // ASP throws error? or clamps?
		if start >= len(s) {
			return "", true
		}
		end := start + length
		if end > len(s) {
			end = len(s)
		}
		return s[start:end], true
	case "instr":
		// InStr([start, ]string1, string2)
		// We implement InStr(s1, s2) and InStr(start, s1, s2)
		var s1, s2 string
		start := 0
		if len(args) == 2 {
			s1 = getStr(0)
			s2 = getStr(1)
		} else if len(args) == 3 {
			start = getInt(0) - 1
			s1 = getStr(1)
			s2 = getStr(2)
		} else {
			return 0, true
		}
		if start < 0 {
			start = 0
		}
		if start >= len(s1) {
			return 0, true
		}
		idx := strings.Index(strings.ToLower(s1[start:]), strings.ToLower(s2))
		if idx == -1 {
			return 0, true
		}
		return idx + start + 1, true // 1-based
	case "replace":
		if len(args) < 3 {
			return "", true
		}
		s := getStr(0)
		find := getStr(1)
		repl := getStr(2)
		// ASP Replace is case-insensitive by default? No, binary default. But usually people want text.
		// strings.ReplaceAll is case sensitive.
		// Implementing case-insensitive replace is harder. Let's stick to ReplaceAll (Case Sensitive)
		// or use strings.Replace with logic.
		// Standard ASP Replace(.., .., .., 1) for text compare.
		return strings.ReplaceAll(s, find, repl), true
	case "trim":
		return strings.TrimSpace(getStr(0)), true
	case "lcase":
		return strings.ToLower(getStr(0)), true
	case "ucase":
		return strings.ToUpper(getStr(0)), true

	// Date/Time
	case "now":
		return time.Now().Format("01/02/2006 15:04:05"), true
	case "date":
		return time.Now().Format("01/02/2006"), true
	case "time":
		return time.Now().Format("15:04:05"), true
	case "year":
		return time.Now().Year(), true
	case "month":
		return int(time.Now().Month()), true
	case "day":
		return time.Now().Day(), true
	case "hour":
		return time.Now().Hour(), true
	case "minute":
		return time.Now().Minute(), true
	case "second":
		return time.Now().Second(), true
	case "weekdayname":
		w := getInt(0)
		// 1=Sunday, 7=Saturday (VBScript default)
		// Go: 0=Sunday
		if w < 1 || w > 7 {
			return "", true
		}
		return time.Weekday((w - 1) % 7).String(), true
	case "monthname":
		m := getInt(0)
		if m < 1 || m > 12 {
			return "", true
		}
		return time.Month(m).String(), true

	// Math
	case "rnd":
		return rand.Float32(), true
	case "round":
		f := getFloat(0)
		return math.Round(f), true
	case "int":
		return int(getFloat(0)), true
	case "abs":
		return math.Abs(getFloat(0)), true
	case "sqr":
		return math.Sqrt(getFloat(0)), true

	// Type Conversion / Misc
	case "cint":
		return getInt(0), true
	case "cstr":
		return getStr(0), true
	case "cdbl":
		return getFloat(0), true
	case "cbool":
		v := getStr(0)
		if strings.ToLower(v) == "true" || v == "1" {
			return true, true
		}
		return false, true
	case "cdate":
		// Simple parser
		t, err := time.Parse("01/02/2006", getStr(0))
		if err == nil {
			return t.Format("01/02/2006"), true
		}
		return getStr(0), true
	case "isnumeric":
		_, err := strconv.ParseFloat(getStr(0), 64)
		return err == nil, true
	case "isdate":
		_, err := time.Parse("01/02/2006", getStr(0))
		return err == nil, true
	case "isempty":
		v := EvaluateExpression(args[0], ctx)
		return v == nil || v == "", true
	case "isnull":
		v := EvaluateExpression(args[0], ctx)
		return v == nil, true

	// Additional Date/Time Functions
	case "datevalue":
		// Parse date from string
		t, err := time.Parse("01/02/2006", getStr(0))
		if err == nil {
			return t.Format("01/02/2006"), true
		}
		return getStr(0), true
	case "timevalue":
		// Parse time from string
		t, err := time.Parse("15:04:05", getStr(0))
		if err == nil {
			return t.Format("15:04:05"), true
		}
		return getStr(0), true
	case "timer":
		// Return seconds since midnight
		now := time.Now()
		midnight := time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, now.Location())
		return int(now.Sub(midnight).Seconds()), true
	case "weekday":
		// Get weekday number (1=Sunday, 7=Saturday)
		if len(args) == 0 {
			return int(time.Now().Weekday()) + 1, true
		}
		t, err := time.Parse("01/02/2006", getStr(0))
		if err == nil {
			return int(t.Weekday()) + 1, true
		}
		return 0, true
	case "dateadd":
		// DateAdd(interval, number, date)
		if len(args) < 3 {
			return getStr(0), true
		}
		interval := strings.ToLower(getStr(0))
		num := getInt(1)
		dateStr := getStr(2)
		t, err := time.Parse("01/02/2006", dateStr)
		if err != nil {
			t = time.Now()
		}
		switch interval {
		case "yyyy":
			t = t.AddDate(num, 0, 0)
		case "m":
			t = t.AddDate(0, num, 0)
		case "d":
			t = t.AddDate(0, 0, num)
		case "h":
			t = t.Add(time.Hour * time.Duration(num))
		case "n":
			t = t.Add(time.Minute * time.Duration(num))
		case "s":
			t = t.Add(time.Second * time.Duration(num))
		}
		return t.Format("01/02/2006"), true
	case "datediff":
		// DateDiff(interval, date1, date2) - difference between two dates
		if len(args) < 3 {
			return 0, true
		}
		interval := strings.ToLower(getStr(0))
		t1Str := getStr(1)
		t2Str := getStr(2)
		t1, err1 := time.Parse("01/02/2006", t1Str)
		t2, err2 := time.Parse("01/02/2006", t2Str)
		if err1 != nil || err2 != nil {
			return 0, true
		}
		diff := t2.Sub(t1)
		switch interval {
		case "d":
			return int(diff.Hours() / 24), true
		case "h":
			return int(diff.Hours()), true
		case "n":
			return int(diff.Minutes()), true
		case "s":
			return int(diff.Seconds()), true
		}
		return 0, true
	case "datepart":
		// DatePart(interval, date)
		if len(args) < 2 {
			return 0, true
		}
		interval := strings.ToLower(getStr(0))
		dateStr := getStr(1)
		t, err := time.Parse("01/02/2006", dateStr)
		if err != nil {
			t = time.Now()
		}
		switch interval {
		case "yyyy":
			return t.Year(), true
		case "m":
			return int(t.Month()), true
		case "d":
			return t.Day(), true
		case "h":
			return t.Hour(), true
		case "n":
			return t.Minute(), true
		case "s":
			return t.Second(), true
		}
		return 0, true
	case "dateserial":
		// DateSerial(year, month, day)
		if len(args) < 3 {
			return "", true
		}
		year := getInt(0)
		month := getInt(1)
		day := getInt(2)
		t := time.Date(year, time.Month(month), day, 0, 0, 0, 0, time.Local)
		return t.Format("01/02/2006"), true
	case "timeserial":
		// TimeSerial(hour, minute, second)
		if len(args) < 3 {
			return "", true
		}
		hour := getInt(0)
		minute := getInt(1)
		second := getInt(2)
		t := time.Date(1900, 1, 1, hour, minute, second, 0, time.Local)
		return t.Format("15:04:05"), true
	case "formatdatetime":
		// FormatDateTime(date, format)
		if len(args) < 2 {
			return getStr(0), true
		}
		dateStr := getStr(0)
		formatType := getInt(1)
		t, err := time.Parse("01/02/2006", dateStr)
		if err != nil {
			return dateStr, true
		}
		switch formatType {
		case 0: // vbGeneralDate (default)
			return t.Format("01/02/2006 15:04:05"), true
		case 1: // vbLongDate
			return t.Format("Monday, January 2, 2006"), true
		case 2: // vbShortDate
			return t.Format("01/02/2006"), true
		case 3: // vbLongTime
			return t.Format("15:04:05"), true
		case 4: // vbShortTime
			return t.Format("15:04"), true
		}
		return dateStr, true

	// Additional String Functions
	case "ltrim":
		return strings.TrimLeft(getStr(0), " \t"), true
	case "rtrim":
		return strings.TrimRight(getStr(0), " \t"), true
	case "space":
		n := getInt(0)
		if n < 0 {
			n = 0
		}
		return strings.Repeat(" ", n), true
	case "string":
		// String(number, character) - returns character repeated
		if len(args) < 2 {
			return "", true
		}
		num := getInt(0)
		char := getStr(1)
		if len(char) > 0 {
			return strings.Repeat(string(char[0]), num), true
		}
		return "", true
	case "instrrev":
		// InStrRev(string, substring, [start], [compare])
		var s1, s2 string
		start := -1
		if len(args) < 2 {
			return 0, true
		}
		s1 = getStr(0)
		s2 = getStr(1)
		if len(args) >= 3 {
			start = getInt(2) - 1 // Convert to 0-based
		}
		if start == -1 {
			start = len(s1) - 1
		}
		if start < 0 || start >= len(s1) {
			return 0, true
		}
		idx := strings.LastIndex(strings.ToLower(s1[:start+1]), strings.ToLower(s2))
		if idx == -1 {
			return 0, true
		}
		return idx + 1, true // Return 1-based index
	case "strreverse":
		s := getStr(0)
		runes := []rune(s)
		for i, j := 0, len(runes)-1; i < j; i, j = i+1, j-1 {
			runes[i], runes[j] = runes[j], runes[i]
		}
		return string(runes), true
	case "strcomp":
		// StrComp(string1, string2, [compare]) - case sensitive comparison
		if len(args) < 2 {
			return 0, true
		}
		s1 := getStr(0)
		s2 := getStr(1)
		caseInsensitive := false
		if len(args) >= 3 {
			caseInsensitive = getInt(2) == 1
		}
		if caseInsensitive {
			s1 = strings.ToLower(s1)
			s2 = strings.ToLower(s2)
		}
		if s1 < s2 {
			return -1, true
		} else if s1 > s2 {
			return 1, true
		}
		return 0, true

	// Additional Math Functions
	case "sin":
		return math.Sin(getFloat(0)), true
	case "cos":
		return math.Cos(getFloat(0)), true
	case "tan":
		return math.Tan(getFloat(0)), true
	case "atn":
		// ArcTangent
		return math.Atan(getFloat(0)), true
	case "log":
		return math.Log(getFloat(0)), true
	case "exp":
		return math.Exp(getFloat(0)), true
	case "sgn":
		// Sign function: -1 if negative, 0 if zero, 1 if positive
		f := getFloat(0)
		if f < 0 {
			return -1, true
		} else if f > 0 {
			return 1, true
		}
		return 0, true
	case "fix":
		// Fix truncates towards zero
		f := getFloat(0)
		if f >= 0 {
			return int(f), true
		}
		return int(math.Ceil(f)), true

	// Conversion Functions (additional)
	case "asc":
		// ASCII value of first character
		s := getStr(0)
		if len(s) > 0 {
			return int(s[0]), true
		}
		return 0, true
	case "ascw":
		// Unicode value of first character
		s := getStr(0)
		if len(s) > 0 {
			return int([]rune(s)[0]), true
		}
		return 0, true
	case "chr":
		// Character from ASCII code
		code := getInt(0)
		if code >= 0 && code <= 255 {
			return string(rune(code)), true
		}
		return "", true
	case "chrw":
		// Character from Unicode code
		code := getInt(0)
		return string(rune(code)), true
	case "hex":
		// Convert to hexadecimal string
		num := getInt(0)
		return fmt.Sprintf("%X", num), true
	case "oct":
		// Convert to octal string
		num := getInt(0)
		return fmt.Sprintf("%o", num), true
	case "cbyte":
		// Byte conversion (treat as int in range 0-255)
		return getInt(0) % 256, true
	case "ccur":
		// Currency conversion (treat as float)
		return getFloat(0), true
	case "clng":
		// Long integer (treat as int in Go)
		return getInt(0), true
	case "csng":
		// Single precision float (treat as float)
		return getFloat(0), true

	// Format Functions
	case "formatcurrency":
		// FormatCurrency(value, [digits], [include_leading_digit], [use_parens], [group_separator])
		value := getFloat(0)
		digits := 2
		if len(args) > 1 {
			digits = getInt(1)
		}
		return fmt.Sprintf("$%."+fmt.Sprintf("%d", digits)+"f", value), true
	case "formatnumber":
		// FormatNumber(value, [digits], [include_leading_digit], [use_parens], [group_separator])
		value := getFloat(0)
		digits := 0
		if len(args) > 1 {
			digits = getInt(1)
		}
		return fmt.Sprintf("%."+fmt.Sprintf("%d", digits)+"f", value), true
	case "formatpercent":
		// FormatPercent(value, [digits], [include_leading_digit], [use_parens], [group_separator])
		value := getFloat(0) * 100
		digits := 0
		if len(args) > 1 {
			digits = getInt(1)
		}
		return fmt.Sprintf("%."+fmt.Sprintf("%d", digits)+"f%%", value), true

	// Array Functions
	case "array":
		// Array(elem1, elem2, ...) - returns array/slice
		// For simplicity, return a slice of interface{}
		var arr []interface{}
		for _, arg := range args {
			arr = append(arr, EvaluateExpression(arg, ctx))
		}
		return arr, true
	case "isarray":
		// Check if variable is array
		v := EvaluateExpression(args[0], ctx)
		_, ok := v.([]interface{})
		return ok, true
	case "split":
		// Split(expression, delimiter, [limit], [compare])
		if len(args) < 2 {
			return nil, true
		}
		expr := getStr(0)
		delim := getStr(1)
		parts := strings.Split(expr, delim)
		var result []interface{}
		for _, p := range parts {
			result = append(result, p)
		}
		return result, true
	case "join":
		// Join(array, delimiter)
		if len(args) < 2 {
			return "", true
		}
		// Get the array
		arr := EvaluateExpression(args[0], ctx)
		delim := getStr(1)
		switch v := arr.(type) {
		case []interface{}:
			var strs []string
			for _, item := range v {
				strs = append(strs, fmt.Sprintf("%v", item))
			}
			return strings.Join(strs, delim), true
		}
		return "", true
	case "lbound":
		// LBound(array, [dimension]) - returns lower bound (always 0 in Go)
		return 0, true
	case "ubound":
		// UBound(array, [dimension]) - returns upper bound
		if len(args) == 0 {
			return 0, true
		}
		arr := EvaluateExpression(args[0], ctx)
		dim := 1
		if len(args) >= 2 {
			dim = getInt(1)
			if dim < 1 {
				dim = 1
			}
		}
		// Traverse nested arrays for requested dimension
		current := arr
		for d := 1; d <= dim; d++ {
			switch v := current.(type) {
			case []interface{}:
				if d == dim {
					return len(v) - 1, true
				}
				if len(v) > 0 {
					current = v[0]
					continue
				}
				// Empty inner dimension
				return -1, true
			default:
				return -1, true
			}
		}
		return -1, true
	case "filter":
		// Filter(array, match, [include], [compare])
		if len(args) < 2 {
			return nil, true
		}
		arrVal := EvaluateExpression(args[0], ctx)
		match := getStr(1)
		include := true
		if len(args) >= 3 {
			include = getInt(2) != 0
		}
		var result []interface{}
		switch v := arrVal.(type) {
		case []interface{}:
			for _, item := range v {
				itemStr := fmt.Sprintf("%v", item)
				hasMatch := strings.Contains(strings.ToLower(itemStr), strings.ToLower(match))
				if hasMatch == include {
					result = append(result, item)
				}
			}
		}
		return result, true

	// Other Functions
	case "scriptengine":
		// Always return "VBScript" for compatibility
		return "VBScript", true
	case "scriptenginebuildversion":
		// Return VBScript build version
		return 18702, true
	case "scriptenginemajorversion":
		// Return VBScript major version (5 = VBScript 5.0)
		return 5, true
	case "scriptengineminorversion":
		// Return VBScript minor version
		return 8, true
	case "typename":
		// Return type name of a variable
		if len(args) == 0 {
			return "Empty", true
		}
		val := EvaluateExpression(args[0], ctx)
		if val == nil {
			return "Empty", true // Was Null, but nil is usually Empty in this engine
		}
		switch val.(type) {
		case bool:
			return "Boolean", true
		case int:
			return "Integer", true
		case float64:
			return "Double", true
		case string:
			return "String", true
		case []interface{}:
			return "Variant()", true
		case Component:
			return "Object", true
		default:
			return "Object", true
		}
	case "vartype":
		// Return variable type constant
		// VBScript VarType constants:
		// 0 = Empty, 1 = Null, 2 = Integer, 3 = Long, 4 = Single, 5 = Double,
		// 8 = String, 9 = Object, 11 = Boolean, 12 = Variant, 13 = DataObject
		if len(args) == 0 {
			return 0, true // Empty
		}
		val := EvaluateExpression(args[0], ctx)
		if val == nil {
			return 0, true // Empty (was Null)
		}
		if val == "" {
			return 8, true // String
		}
		switch val.(type) {
		case bool:
			return 11, true // Boolean
		case int:
			return 2, true // Integer
		case float64:
			return 5, true // Double
		case string:
			return 8, true // String
		case []interface{}:
			// Array of Variant: base type vbVariant (12) plus vbArray flag (8192) => 8204
			return 8204, true
		case Component:
			return 9, true // Object
		default:
			return 9, true // Object
		}
	case "rgb":
		// RGB(red, green, blue) - returns color as string in hex format
		if len(args) < 3 {
			return "#000000", true
		}
		r := getInt(0) % 256
		g := getInt(1) % 256
		b := getInt(2) % 256
		if r < 0 {
			r = 0
		}
		if g < 0 {
			g = 0
		}
		if b < 0 {
			b = 0
		}
		return fmt.Sprintf("#%02X%02X%02X", r, g, b), true
	case "isobject":
		// Check if variable is an object (in our case, non-primitive types)
		if len(args) == 0 {
			return false, true
		}
		val := EvaluateExpression(args[0], ctx)
		switch val.(type) {
		case []interface{}, map[string]interface{}, Component:
			return true, true
		default:
			return val == nil && val != "", true
		}
	case "createobject":
		// CreateObject(progid)
		if len(args) == 0 {
			return nil, true
		}
		progid := getStr(0) // Don't ToLower here, ComponentFactory handles it
		return ComponentFactory(progid, ctx), true

	case "eval":
		// Eval(expression) - evaluate expression string and return result
		if len(args) == 0 {
			return nil, true
		}
		exprStr := getStr(0)
		// Evaluate the expression string
		result := EvaluateExpression(exprStr, ctx)
		return result, true
	}

	return nil, false
}

func getFloatVal(v interface{}) float64 {
	if f, ok := v.(float64); ok {
		return f
	}
	if i, ok := v.(int); ok {
		return float64(i)
	}
	if s, ok := v.(string); ok {
		if f, err := strconv.ParseFloat(s, 64); err == nil {
			return f
		}
	}
	return 0.0
}

// parseVBScriptInt parses VBScript numeric literals including hexadecimal (&h) and octal (&o) formats
func parseVBScriptInt(s string) (int, bool) {
	sLower := strings.ToLower(strings.TrimSpace(s))

	// 1. Check for Hexadecimal (&h or &H)
	if strings.HasPrefix(sLower, "&h") {
		// Remove the "&h" prefix and convert base 16
		val, err := strconv.ParseInt(sLower[2:], 16, 64)
		if err == nil {
			return int(val), true
		}
	}

	// 2. Check for Octal (&o or &O)
	if strings.HasPrefix(sLower, "&o") {
		// Remove the "&o" prefix and convert base 8
		val, err := strconv.ParseInt(sLower[2:], 8, 64)
		if err == nil {
			return int(val), true
		}
	}

	// 3. Try standard decimal
	val, err := strconv.Atoi(s)
	if err == nil {
		return val, true
	}

	return 0, false
}
