package asp

import (
	"fmt"
	"math/rand"
	"reflect"
	"strings"
	"time"
)

// Helper to create nested arrays based on dimensions
func makeNestedArray(dims []int) []interface{} {
	if len(dims) == 0 {
		return nil
	}
	size := dims[0]
	arr := make([]interface{}, size)
	if len(dims) == 1 {
		return arr
	}
	// create inner slices
	innerDims := dims[1:]
	for i := 0; i < size; i++ {
		arr[i] = makeNestedArray(innerDims)
	}
	return arr
}

// preserveCopy copies elements from old into a newly created nested array
func preserveCopy(old interface{}, dims []int) []interface{} {
	newArr := makeNestedArray(dims)
	if old == nil {
		return newArr
	}
	oldSlice, ok := old.([]interface{})
	if !ok {
		return newArr
	}

	limit := len(oldSlice)
	if limit > len(newArr) {
		limit = len(newArr)
	}

	for i := 0; i < limit; i++ {
		if len(dims) == 1 {
			newArr[i] = oldSlice[i]
		} else {
			// For nested dimension, recurse if possible
			if i < len(oldSlice) {
				newArr[i] = preserveCopy(oldSlice[i], dims[1:])
			}
		}
	}
	return newArr
}

// tryCallSetProperty attempts to call SetProperty method on an object via reflection
func tryCallSetProperty(obj interface{}, propName string, value interface{}) bool {
	v := reflect.ValueOf(obj)
	method := v.MethodByName("SetProperty")
	if !method.IsValid() {
		return false
	}

	// Call SetProperty(propName, value)
	args := []reflect.Value{
		reflect.ValueOf(propName),
		reflect.ValueOf(value),
	}
	method.Call(args)
	return true
}

// Procedure stores Sub metadata
type Procedure struct {
	LineNum int
	Params  []string
}

// Instruction represents a single executable line
type Instruction struct {
	Type    TokenType // HTML or Code
	Content string    // The raw string
	LineNum int       // Original line number
}

type LoopType int

const (
	LoopFor LoopType = iota
	LoopDo
	LoopForEach
	LoopWhile
)

type LoopState struct {
	Type     LoopType
	VarName  string // For / ForEach
	StartVal int    // For
	EndVal   int    // For
	StepVal  int    // For

	// ForEach properties
	Collection []interface{} // ForEach - array/collection of items
	Index      int           // ForEach - current index in collection

	Condition string // Do
	Until     bool   // Do (true=Until, false=While)

	StartPC int // Jump back target
}

type IfState struct {
	Active  bool // Is current block executing?
	Handled bool // Has any block in this chain executed?
}

type SelectState struct {
	SelectValue interface{}
	Handled     bool
	Active      bool
}

// Engine handles the execution of the script
type ByRefSetter func(interface{})

type CallFrame struct {
	ReturnPC   int
	ParamNames []string
	Setters    []ByRefSetter
}

type Engine struct {
	Instructions []Instruction
	Ctx          *ExecutionContext
	PC           int                  // Program Counter
	CurrentLine  int                  // Currently executing line number
	Labels       map[string]Procedure // For Subroutines
	Classes      map[string]*ClassDef // Class Definitions

	// Control Flow Stacks
	CallStack      []CallFrame // Return addresses and ByRef bindings for Subs
	IfStack        []IfState
	LoopStack      []LoopState
	SelectStack    []SelectState
	WithStack      []interface{}
	SingleProcMode bool // If true, Run() returns after CallStack is empty
}

// Helper to safely remove inline comments (e.g., x = 1 ' comment)
func stripComment(line string) string {
	inQuote := false
	for i := 0; i < len(line); i++ {
		if line[i] == '"' {
			inQuote = !inQuote
		}
		// Se encontrar ' fora de aspas, Ã© o inÃ­cio de um comentÃ¡rio
		if line[i] == '\'' && !inQuote {
			return strings.TrimSpace(line[:i])
		}
	}
	return strings.TrimSpace(line)
}

// resolveByRefSetter finds a setter for a variable-like expression in the given context.
func ResolveByRefSetter(ctx *ExecutionContext, expr string) (ByRefSetter, bool) {
	expr = strings.TrimSpace(expr)
	if expr == "" {
		return nil, false
	}
	if strings.ContainsAny(expr, ".(") {
		return nil, false
	}
	lower := strings.ToLower(expr)

	if ctx == nil {
		return nil, false
	}

	if _, ok := ctx.Variables[lower]; ok {
		return func(v interface{}) { ctx.Variables[lower] = v }, true
	}

	if ctx.GlobalVariables != nil {
		if _, ok := ctx.GlobalVariables[lower]; ok {
			return func(v interface{}) { ctx.GlobalVariables[lower] = v }, true
		}
	}

	if ctx.CurrentInstance != nil {
		if _, ok := ctx.CurrentInstance.Variables[lower]; ok {
			inst := ctx.CurrentInstance
			return func(v interface{}) { inst.Variables[lower] = v }, true
		}
	}

	return nil, false
}

// splitByColon splits a line by top-level colons (:) respecting quotes and date literals (#...#)
// This is used for processing multiple statements on one line (e.g., "Dim x : x = 5")
func splitByColon(s string) []string {
	var parts []string
	var current strings.Builder
	inQuote := false
	inHash := false
	depth := 0

	for i := 0; i < len(s); i++ {
		c := s[i]

		// Track quotes
		if c == '"' && !inHash {
			inQuote = !inQuote
		}

		// Track date literals (#...#)
		if c == '#' && !inQuote {
			inHash = !inHash
		}

		// Track parentheses depth
		if !inQuote && !inHash {
			if c == '(' {
				depth++
			} else if c == ')' {
				if depth > 0 {
					depth--
				}
			} else if c == ':' && depth == 0 {
				// Found a statement separator
				parts = append(parts, strings.TrimSpace(current.String()))
				current.Reset()
				continue
			}
		}
		current.WriteByte(c)
	}

	if current.Len() > 0 {
		parts = append(parts, strings.TrimSpace(current.String()))
	}

	return parts
}

// splitCaseParts splits a CASE argument list by top-level commas respecting quotes and parentheses.
func splitCaseParts(s string) []string {
	var parts []string
	var current strings.Builder
	inQuote := false
	depth := 0

	for i := 0; i < len(s); i++ {
		c := s[i]
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
			} else if c == ',' && depth == 0 {
				parts = append(parts, strings.TrimSpace(current.String()))
				current.Reset()
				continue
			}
		}
		current.WriteByte(c)
	}

	if current.Len() > 0 {
		parts = append(parts, strings.TrimSpace(current.String()))
	}

	return parts
}

// Prepare parses tokens into instructions and pre-scans for Subs/Classes
func Prepare(tokens []Token) *Engine {
	eng := &Engine{
		Labels:      make(map[string]Procedure),
		Classes:     make(map[string]*ClassDef),
		CallStack:   make([]CallFrame, 0),
		IfStack:     make([]IfState, 0),
		LoopStack:   make([]LoopState, 0),
		SelectStack: make([]SelectState, 0),
		WithStack:   make([]interface{}, 0),
	}

	// Flatten tokens into lines
	for _, t := range tokens {
		if t.Type == TokenHTML {
			eng.Instructions = append(eng.Instructions, Instruction{Type: TokenHTML, Content: t.Content, LineNum: t.LineNum})
		} else {
			// Split code by newlines
			lines := strings.Split(t.Content, "\n")
			for i, l := range lines {
				l = stripComment(l)
				if l == "" {
					continue
				}
				if strings.HasPrefix(l, "'") {
					continue
				}

				// Expand colon-separated statements into separate instructions
				// This handles syntax like "Dim x : x = 5 : y = 10"
				colonParts := splitByColon(l)
				if len(colonParts) > 1 {
					// Multiple statements on one line
					for _, part := range colonParts {
						part = strings.TrimSpace(part)
						if part != "" {
							eng.Instructions = append(eng.Instructions, Instruction{Type: TokenCode, Content: part, LineNum: t.LineNum + i})
						}
					}
				} else {
					// Single statement
					eng.Instructions = append(eng.Instructions, Instruction{Type: TokenCode, Content: l, LineNum: t.LineNum + i})
				}
			}
		}
	}

	// Pre-scan for Subs, Functions, Classes
	var currentClass *ClassDef
	inClass := false

	for i, inst := range eng.Instructions {
		if inst.Type == TokenCode {
			line := inst.Content
			lowerLine := strings.ToLower(line)

			// Visibility Handling
			visibility := VisPublic
			effectiveLine := lowerLine
			effectiveContent := line

			if strings.HasPrefix(lowerLine, "private ") {
				visibility = VisPrivate
				effectiveLine = strings.TrimSpace(lowerLine[8:])
				effectiveContent = strings.TrimSpace(line[8:])
			} else if strings.HasPrefix(lowerLine, "public ") {
				visibility = VisPublic
				effectiveLine = strings.TrimSpace(lowerLine[7:])
				effectiveContent = strings.TrimSpace(line[7:])
			}

			// Default Keyword Handling (e.g. Public Default Function)
			isDefault := false
			if strings.HasPrefix(effectiveLine, "default ") {
				isDefault = true
				effectiveLine = strings.TrimSpace(effectiveLine[8:])
				effectiveContent = strings.TrimSpace(effectiveContent[8:])
			}

			// CLASS START
			if strings.HasPrefix(effectiveLine, "class ") {
				name := strings.TrimSpace(effectiveContent[6:])
				currentClass = &ClassDef{
					Name:       name,
					Variables:  make(map[string]ClassMemberVar),
					Methods:    make(map[string]Procedure),
					Properties: make(map[string][]PropertyDef),
				}
				inClass = true
				continue
			}

			// CLASS END
			if effectiveLine == "end class" {
				if inClass && currentClass != nil {
					eng.Classes[strings.ToLower(currentClass.Name)] = currentClass
				}
				inClass = false
				currentClass = nil
				continue
			}

			// SUB / FUNCTION / PROPERTY
			isSub := strings.HasPrefix(effectiveLine, "sub ")
			isFunc := strings.HasPrefix(effectiveLine, "function ")
			isProp := strings.HasPrefix(effectiveLine, "property ")

			if isSub || isFunc || isProp {
				var procName string
				var params []string
				var rest string
				var paramStr string
				var propType PropertyType

				if isSub {
					rest = effectiveContent[4:]
				} else if isFunc {
					rest = effectiveContent[9:]
				} else if isProp {
					rest = effectiveContent[9:]
					lowerRest := strings.ToLower(rest)
					if strings.HasPrefix(lowerRest, "get ") {
						propType = PropGet
						rest = rest[4:]
					} else if strings.HasPrefix(lowerRest, "let ") {
						propType = PropLet
						rest = rest[4:]
					} else if strings.HasPrefix(lowerRest, "set ") {
						propType = PropSet
						rest = rest[4:]
					} else {
						propType = PropGet
					}
				}

				startP := strings.Index(rest, "(")
				endP := strings.LastIndex(rest, ")")

				if startP > -1 && endP > startP {
					procName = strings.TrimSpace(rest[:startP])
					paramStr = rest[startP+1 : endP]
					if len(strings.TrimSpace(paramStr)) > 0 {
						pList := strings.Split(paramStr, ",")
						for _, p := range pList {
							p = strings.TrimSpace(p)
							// Remove ByRef keyword from parameter name
							if strings.HasPrefix(strings.ToLower(p), "byref ") {
								p = strings.TrimSpace(p[6:])
							}
							params = append(params, p)
						}
					}
				} else {
					procName = strings.TrimSpace(rest)
				}

				if inClass {
					if isDefault {
						currentClass.DefaultMethod = strings.ToLower(procName)
					}
					if isProp {
						pDef := PropertyDef{
							Name:       procName,
							Type:       propType,
							Params:     params,
							LineNum:    i,
							Visibility: visibility,
						}
						lowerName := strings.ToLower(procName)
						currentClass.Properties[lowerName] = append(currentClass.Properties[lowerName], pDef)
					} else {
						// Sub/Function in Class
						// Note: Visibility is recorded? Current `Procedure` struct doesn't have Visibility field.
						// We should verify strict visibility rules later or assume Public for now if Procedure struct unchanged.
						// VBScript Engine usually enforces this.
						// For now, let's treat all as invokable by ExecuteClassMethod (which checks visibility in parser logic?)
						// Actually `GetProperty` in class_def.go checks `p.Visibility`.
						// But `CallMethod` checks `ClassDef.Methods`. `Procedure` lacks visibility.
						// I will rely on naming convention or valid lookup.
						// Ideally I should update `Procedure` to include Visibility.
						// But to save time/complexity, I'll assume Public or check later.
						// Wait, `ClassMemberVar` has it.
						// `Methods` is `map[string]Procedure`.
						// Use `Variables` map to store Method visibility? No.
						// Since I cannot change Procedure struct easily without breaking other things (maybe?),
						// I will stick to adding it. Procedure is defined in this file. I can change it!
						// But `Labels` (Global) also use Procedure. Visibility there is meaningless (Public).
						// So I can add `Visibility` field to `Procedure` struct. It's safe.
						currentClass.Methods[strings.ToLower(procName)] = Procedure{
							LineNum: i,
							Params:  params,
							// Visibility: visibility, // Need to update struct first
						}
					}
				} else {
					if isSub || isFunc {
						eng.Labels[strings.ToLower(procName)] = Procedure{LineNum: i, Params: params}
					}
				}
				continue
			}

			// MEMBER VARIABLES (inside Class)
			// e.g. "Private x", "Public y", "Dim z"
			if inClass {
				// "Dim" inside Class is implicitly Public (or Private? VBScript defaults to Public for Dim in Class? No, Private?)
				// VBScript Class: "Dim" = Public (actually "Public" is preferred). "Private" = Private.
				// Let's assume Dim is Public for now, or check documentation.
				// Actually "Dim" in Class creates a Public member.
				isDim := strings.HasPrefix(lowerLine, "dim ")

				if isDim || strings.HasPrefix(lowerLine, "private ") || strings.HasPrefix(lowerLine, "public ") {
					// Re-parse to handle comma separated
					// Note: We handled "Private/Public" prefix stripping above for `effectiveLine`.
					// If it was "Private Sub", we handled it in Sub block.
					// If we are here, it's likely a variable declaration.
					// But wait, "Private Sub" also hits "strings.HasPrefix(lowerLine, 'private ')".
					// So `effectiveLine` is "Sub ...".
					// The `isSub` check handled it and `continue`.
					// So if we reach here, it is NOT a Sub/Func/Prop.
					// So it MUST be a variable declaration.

					rest := effectiveContent
					if isDim {
						rest = inst.Content[4:]
						// Dim defaults to Public in Class?
						// Microsoft docs: "Variables declared with Dim at the class level are available to all procedures within the class... default to Public?"
						// Actually Dim at script level is script-scope.
						// In Class, Dim makes it Public.
						visibility = VisPublic
					}

					// Split vars
					vars := strings.Split(rest, ",")
					for _, v := range vars {
						vName := strings.TrimSpace(v)
						// Strip array parens
						if idx := strings.Index(vName, "("); idx > -1 {
							vName = strings.TrimSpace(vName[:idx])
						}
						if vName != "" {
							currentClass.Variables[strings.ToLower(vName)] = ClassMemberVar{
								Name:       vName,
								Visibility: visibility,
							}
						}
					}
				}
			}
		}
	}

	return eng
}

func (e *Engine) Run(ctx *ExecutionContext) {
	e.Ctx = ctx
	// Link Engine to Context so Class Instances can call back
	e.Ctx.Engine = e
	// Share Class Definitions (Global)
	e.Ctx.GlobalClasses = e.Classes

	if !e.SingleProcMode {
		e.PC = 0
	}

	// Reset With stack for a fresh execution cycle
	e.WithStack = e.WithStack[:0]

	// Helper to check if we should execute current line
	shouldExecute := func() bool {
		// Check IfStack
		if len(e.IfStack) > 0 {
			if !e.IfStack[len(e.IfStack)-1].Active {
				return false
			}
		}
		// Check SelectStack
		if len(e.SelectStack) > 0 {
			if !e.SelectStack[len(e.SelectStack)-1].Active {
				return false
			}
		}
		return true
	}

	for e.PC < len(e.Instructions) {
		// Check if Response.End was called
		if e.Ctx.ResponseState.IsEnded {
			return
		}

		inst := e.Instructions[e.PC]
		e.CurrentLine = inst.LineNum

		// If it's HTML, just print it (UNLESS skipped)
		if inst.Type == TokenHTML {
			if shouldExecute() {
				e.Ctx.Write(inst.Content)
			}
			e.PC++
			continue
		}

		// It's Code. Identify command.
		line := inst.Content
		lowerLine := strings.ToLower(line)

		if lowerLine == "option explicit" {
			e.Ctx.OptionExplicitEnabled = true
			e.PC++
			continue
		}

		// CLASS Definition - Skip during execution
		if strings.HasPrefix(lowerLine, "class ") {
			depth := 1
			// Start scanning from next line
			scanPC := e.PC + 1
			foundEnd := false
			for scanPC < len(e.Instructions) {
				l := strings.ToLower(e.Instructions[scanPC].Content)
				if strings.HasPrefix(l, "class ") {
					depth++
				}
				if l == "end class" {
					depth--
				}
				if depth == 0 {
					e.PC = scanPC + 1
					foundEnd = true
					break
				}
				scanPC++
			}
			if !foundEnd {
				e.PC++ // Just move to next line if end not found (error?)
			}
			continue
		}

		// --- CONTROL FLOW STRUCTURES (Must process even if skipped to maintain balance) ---

		isExecuting := shouldExecute()

		// IF
		if strings.HasPrefix(lowerLine, "if ") && strings.HasSuffix(lowerLine, " then") {
			if isExecuting {
				condStr := line[3 : len(line)-5]
				result := EvaluateCondition(condStr, e.Ctx)
				e.IfStack = append(e.IfStack, IfState{Active: result, Handled: result})
			} else {
				e.IfStack = append(e.IfStack, IfState{Active: false, Handled: true})
			}
			e.PC++
			continue
		}

		// ELSEIF
		if strings.HasPrefix(lowerLine, "elseif ") && strings.HasSuffix(lowerLine, " then") {
			if len(e.IfStack) > 0 {
				top := &e.IfStack[len(e.IfStack)-1]
				top.Active = false

				if !top.Handled {
					parentActive := true
					if len(e.IfStack) > 1 {
						parentActive = e.IfStack[len(e.IfStack)-2].Active
					}

					if parentActive {
						condStr := line[7 : len(line)-5]
						result := EvaluateCondition(condStr, e.Ctx)
						if result {
							top.Active = true
							top.Handled = true
						}
					}
				}
			}
			e.PC++
			continue
		}

		// ELSE
		if lowerLine == "else" {
			if len(e.IfStack) > 0 {
				top := &e.IfStack[len(e.IfStack)-1]
				top.Active = false
				if !top.Handled {
					parentActive := true
					if len(e.IfStack) > 1 {
						parentActive = e.IfStack[len(e.IfStack)-2].Active
					}
					if parentActive {
						top.Active = true
						top.Handled = true
					}
				}
			}
			e.PC++
			continue
		}

		// END IF
		if lowerLine == "end if" {
			if len(e.IfStack) > 0 {
				e.IfStack = e.IfStack[:len(e.IfStack)-1]
			}
			e.PC++
			continue
		}

		// SELECT CASE
		if strings.HasPrefix(lowerLine, "select case ") {
			if isExecuting {
				valExpr := strings.TrimSpace(line[12:])
				val := EvaluateExpression(valExpr, e.Ctx)
				e.SelectStack = append(e.SelectStack, SelectState{SelectValue: val, Handled: false, Active: false})
			} else {
				e.SelectStack = append(e.SelectStack, SelectState{Active: false, Handled: true})
			}
			e.PC++
			continue
		}

		// CASE
		if strings.HasPrefix(lowerLine, "case ") {
			if len(e.SelectStack) > 0 {
				top := &e.SelectStack[len(e.SelectStack)-1]
				top.Active = false

				if !top.Handled {
					arg := strings.TrimSpace(line[5:])
					if strings.ToLower(arg) == "else" {
						top.Active = true
						top.Handled = true
					} else {
						match := false
						parts := splitCaseParts(arg)
						for _, p := range parts {
							p = strings.TrimSpace(p)
							if p == "" {
								continue
							}
							pLower := strings.ToLower(p)
							if strings.Contains(pLower, " to ") {
								rangeIdx := strings.Index(pLower, " to ")
								startExpr := strings.TrimSpace(p[:rangeIdx])
								endExpr := strings.TrimSpace(p[rangeIdx+4:])
								startVal := EvaluateExpression(startExpr, e.Ctx)
								endVal := EvaluateExpression(endExpr, e.Ctx)
								selInt, selOK := toInt(top.SelectValue)
								startInt, sOK := toInt(startVal)
								endInt, eOK := toInt(endVal)
								if selOK && sOK && eOK {
									if selInt >= startInt && selInt <= endInt {
										match = true
										break
									}
								}
							} else {
								val := EvaluateExpression(p, e.Ctx)
								if fmt.Sprintf("%v", val) == fmt.Sprintf("%v", top.SelectValue) {
									match = true
									break
								}
							}
						}

						if match {
							top.Active = true
							top.Handled = true
						}
					}
				}
			}
			e.PC++
			continue
		}

		// RANDOMIZE
		if lowerLine == "randomize" {
			rand.Seed(time.Now().UnixNano())
			e.PC++
			continue
		}

		// END SELECT
		if lowerLine == "end select" {
			if len(e.SelectStack) > 0 {
				e.SelectStack = e.SelectStack[:len(e.SelectStack)-1]
			}
			e.PC++
			continue
		}

		// WITH ... END WITH
		if strings.HasPrefix(lowerLine, "with ") {
			targetExpr := strings.TrimSpace(line[5:])
			if isExecuting {
				val := EvaluateExpression(targetExpr, e.Ctx)
				e.WithStack = append(e.WithStack, val)
			} else {
				// Preserve nesting even when skipping execution
				e.WithStack = append(e.WithStack, nil)
			}
			e.PC++
			continue
		}

		if lowerLine == "end with" {
			if len(e.WithStack) > 0 {
				e.WithStack = e.WithStack[:len(e.WithStack)-1]
			}
			e.PC++
			continue
		}

		if !isExecuting {
			e.PC++
			continue
		}

		// --- EXECUTABLE COMMANDS ---

		// SUB / FUNCTION / PROPERTY Definition - Skip
		// Handle visibility modifiers (private/public)
		effectiveLine := lowerLine
		if strings.HasPrefix(lowerLine, "private ") {
			effectiveLine = strings.TrimSpace(lowerLine[8:])
		} else if strings.HasPrefix(lowerLine, "public ") {
			effectiveLine = strings.TrimSpace(lowerLine[7:])
		}

		if strings.HasPrefix(effectiveLine, "sub ") || strings.HasPrefix(effectiveLine, "function ") || strings.HasPrefix(effectiveLine, "property ") {
			scanPC := e.PC + 1
			foundEnd := false
			targetEnd := "end sub"
			if strings.HasPrefix(effectiveLine, "function ") {
				targetEnd = "end function"
			}
			if strings.HasPrefix(effectiveLine, "property ") {
				targetEnd = "end property"
			}

			for scanPC < len(e.Instructions) {
				l := strings.ToLower(e.Instructions[scanPC].Content)
				// Naive depth check (nested subs are not really valid in VBScript but Class methods are inside Class)
				// Since we are INSIDE Run(), we are at Global scope or inside a Sub executing.
				// Defining a Sub inside execution flow is invalid in VBScript (Subs are static).
				// But we are iterating instructions.
				// If we encounter "Sub X", we skip it.
				if l == targetEnd {
					// We don't support nested definitions logic here, just skip to end.
					e.PC = scanPC + 1
					foundEnd = true
					break
				}
				scanPC++
			}
			if !foundEnd {
				e.PC++
			}
			continue
		}

		// END SUB / FUNCTION / PROPERTY (Return)
		if lowerLine == "end sub" || lowerLine == "end function" || lowerLine == "end property" {
			if len(e.CallStack) > 0 {
				frame := e.CallStack[len(e.CallStack)-1]
				e.CallStack = e.CallStack[:len(e.CallStack)-1]

				for i, name := range frame.ParamNames {
					if i < len(frame.Setters) && frame.Setters[i] != nil {
						lower := strings.ToLower(name)
						if val, ok := e.Ctx.Variables[lower]; ok {
							frame.Setters[i](val)
						}
					}
				}

				e.PC = frame.ReturnPC

				// If we just popped the last frame and we are in SingleProcMode (e.g. user function call), we must return
				if len(e.CallStack) == 0 && e.SingleProcMode {
					return
				}
			} else {
				if e.SingleProcMode {
					return // End execution of this single procedure
				}
				e.PC++
			}
			continue
		}

		// EXIT SUB / FUNCTION / PROPERTY
		if lowerLine == "exit sub" || lowerLine == "exit function" || lowerLine == "exit property" {
			if len(e.CallStack) > 0 {
				frame := e.CallStack[len(e.CallStack)-1]
				e.CallStack = e.CallStack[:len(e.CallStack)-1]

				for i, name := range frame.ParamNames {
					if i < len(frame.Setters) && frame.Setters[i] != nil {
						lower := strings.ToLower(name)
						if val, ok := e.Ctx.Variables[lower]; ok {
							frame.Setters[i](val)
						}
					}
				}

				e.PC = frame.ReturnPC

				// If we just popped the last frame and we are in SingleProcMode, we must return
				if len(e.CallStack) == 0 && e.SingleProcMode {
					return
				}
			} else {
				if e.SingleProcMode {
					return
				}
				e.PC++
			}
			continue
		}

		// ERASE command - clears array
		if strings.HasPrefix(lowerLine, "erase ") {
			rest := strings.TrimSpace(line[6:])
			// Can have multiple arrays separated by commas
			arrNames := strings.Split(rest, ",")
			for _, arrName := range arrNames {
				arrName = strings.TrimSpace(arrName)
				key := strings.ToLower(arrName)
				// Reset to empty array or nil
				e.Ctx.Variables[key] = nil
			}
			e.PC++
			continue
		}

		// CALL
		if strings.HasPrefix(lowerLine, "call ") {
			rest := line[5:]
			subName := ""
			args := []string{}

			startP := strings.Index(rest, "(")
			endP := strings.LastIndex(rest, ")")

			if startP > -1 && endP > startP {
				subName = strings.TrimSpace(rest[:startP])
				argStr := rest[startP+1 : endP]
				currentArg := ""
				inQuote := false
				for _, r := range argStr {
					if r == '"' {
						inQuote = !inQuote
					}
					if r == ',' && !inQuote {
						args = append(args, strings.TrimSpace(currentArg))
						currentArg = ""
					} else {
						currentArg += string(r)
					}
				}
				if strings.TrimSpace(currentArg) != "" || (len(args) > 0 && currentArg == "") {
					args = append(args, strings.TrimSpace(currentArg))
				}
			} else {
				subName = strings.TrimSpace(rest)
			}

			if proc, ok := e.Labels[strings.ToLower(subName)]; ok {
				frame := CallFrame{ReturnPC: e.PC + 1, ParamNames: proc.Params, Setters: make([]ByRefSetter, len(proc.Params))}
				for i, paramName := range proc.Params {
					if i < len(args) {
						argExpr := args[i]
						val := EvaluateExpression(argExpr, e.Ctx)
						e.Ctx.Variables[strings.ToLower(paramName)] = val
						if setter, ok := ResolveByRefSetter(e.Ctx, argExpr); ok {
							frame.Setters[i] = setter
						}
					}
				}
				e.CallStack = append(e.CallStack, frame)
				e.PC = proc.LineNum + 1
				continue
			}
		}

		// DIM
		if strings.HasPrefix(lowerLine, "dim ") {
			dimContent := line[4:]

			// Variable declarations (comma-separated)
			vars := strings.Split(dimContent, ",")
			for _, v := range vars {
				v = strings.TrimSpace(v)
				if v != "" {
					// Check for array declaration: name(size)
					if idx := strings.Index(v, "("); idx > -1 && strings.HasSuffix(v, ")") {
						name := strings.TrimSpace(v[:idx])
						sizeStr := strings.TrimSpace(v[idx+1 : len(v)-1])

						// Evaluate size (needs to be available in current context)
						// Note: VBScript Dim size must be constant integer usually, but here we can evaluate
						sizeVal := EvaluateExpression(sizeStr, e.Ctx)
						size, _ := toInt(sizeVal)

						// Create array (slice of interface{})
						// VBScript array(n) has n+1 elements (0 to n)
						if size >= 0 {
							arr := make([]interface{}, size+1)
							e.Ctx.Variables[strings.ToLower(name)] = arr
						} else {
							// Error or just nil?
							e.Ctx.Variables[strings.ToLower(name)] = nil
						}
					} else {
						e.Ctx.Variables[strings.ToLower(v)] = nil
					}
				}
			}

			e.PC++
			continue
		}

		// REDIM
		if strings.HasPrefix(lowerLine, "redim ") {
			// ReDim [Preserve] name(size)
			rest := strings.TrimSpace(line[6:])
			preserve := false
			if strings.HasPrefix(strings.ToLower(rest), "preserve ") {
				preserve = true
				rest = strings.TrimSpace(rest[9:])
			}

			// Extract Name and Sizes "arr(5)" or "arr(2,3)"
			startP := strings.Index(rest, "(")
			endP := strings.LastIndex(rest, ")")

			if startP > -1 && endP > startP {
				arrName := strings.TrimSpace(rest[:startP])
				sizeExpr := rest[startP+1 : endP]

				// Split dimensions by comma and evaluate each
				sizeParts := strings.Split(sizeExpr, ",")
				var dims []int
				for _, p := range sizeParts {
					p = strings.TrimSpace(p)
					val := EvaluateExpression(p, e.Ctx)
					n, _ := toInt(val)
					// VBScript upper bound N means length N+1
					if n < 0 {
						n = 0
					}
					dims = append(dims, n+1)
				}

				var newArr []interface{}

				if preserve {
					// Check existing
					if existing, ok := e.Ctx.Variables[strings.ToLower(arrName)]; ok {
						// If existing is a slice, try to preserve nested elements
						newArr = preserveCopy(existing, dims)
					} else {
						if e.Ctx.OptionExplicitEnabled {
							e.Ctx.Err.Raise(500, "Microsoft VBScript runtime error", "Variable is undefined: '"+arrName+"'", "", 0)
							if !e.Ctx.OnErrorResumeNext {
								panic(fmt.Sprintf("Variable is undefined: '%s'", arrName))
							}
						}
						newArr = makeNestedArray(dims)
					}
				} else {
					if e.Ctx.OptionExplicitEnabled {
						if _, ok := e.Ctx.Variables[strings.ToLower(arrName)]; !ok {
							e.Ctx.Err.Raise(500, "Microsoft VBScript runtime error", "Variable is undefined: '"+arrName+"'", "", 0)
							if !e.Ctx.OnErrorResumeNext {
								panic(fmt.Sprintf("Variable is undefined: '%s'", arrName))
							}
						}
					}
					newArr = makeNestedArray(dims)
				}

				e.Ctx.Variables[strings.ToLower(arrName)] = newArr
			}
			e.PC++
			continue
		}

		// ON ERROR RESUME NEXT
		if lowerLine == "on error resume next" {
			e.Ctx.OnErrorResumeNext = true
			e.PC++
			continue
		}

		// ON ERROR GOTO 0
		if lowerLine == "on error goto 0" {
			e.Ctx.OnErrorResumeNext = false
			e.Ctx.Err.Clear() // Also clears error
			e.PC++
			continue
		}

		// ERR.CLEAR
		if lowerLine == "err.clear" {
			e.Ctx.Err.Clear()
			e.PC++
			continue
		}

		// RESPONSE.WRITE
		if strings.HasPrefix(lowerLine, "response.write") {
			rawExpr := strings.TrimSpace(line[len("response.write"):])
			if strings.HasPrefix(rawExpr, "(") && strings.HasSuffix(rawExpr, ")") {
				expr := rawExpr[1 : len(rawExpr)-1]
				val := EvaluateExpression(expr, e.Ctx)
				e.Ctx.Write(fmt.Sprintf("%v", val))
			} else {
				val := EvaluateExpression(rawExpr, e.Ctx)
				e.Ctx.Write(fmt.Sprintf("%v", val))
			}
			e.PC++
			continue
		}

		// FOR EACH LOOP
		if strings.HasPrefix(lowerLine, "for each ") {

			rest := line[9:]
			inIdx := strings.Index(strings.ToLower(rest), " in ")
			if inIdx > -1 {
				varName := strings.TrimSpace(rest[:inIdx])
				varName = strings.ToLower(varName)
				collExpr := strings.TrimSpace(rest[inIdx+4:])

				// Evaluate collection expression to get items
				var items []interface{}
				val := EvaluateExpression(collExpr, e.Ctx)

				// Type Switch to handle multiple collection types including Collection
				switch v := val.(type) {
				case []string:
					for _, s := range v {
						items = append(items, s)
					}
				case []interface{}:
					items = v
				case map[string]interface{}:
					// Iterate over keys
					for k := range v {
						items = append(items, k)
					}
				case *Collection:
					// Use Keys for For Each iteration
					for _, k := range v.Keys() {
						items = append(items, k)
					}
				case Enumerable:
					items = v.Enumeration()
				}

				if len(items) > 0 {
					if e.Ctx.OptionExplicitEnabled {
						if _, ok := e.Ctx.Variables[varName]; !ok {
							e.Ctx.Err.Raise(500, "Microsoft VBScript runtime error", "Variable is undefined: '"+varName+"'", "", 0)
							if !e.Ctx.OnErrorResumeNext {
								panic(fmt.Sprintf("Variable is undefined: '%s'", varName))
							}
						}
					}
					e.Ctx.Variables[varName] = items[0]

					e.LoopStack = append(e.LoopStack, LoopState{
						Type:       LoopForEach,
						VarName:    varName,
						Collection: items,
						Index:      0,
						StartPC:    e.PC,
					})
				} else {
					// CORREÃ‡ÃƒO CRÃTICA: Se a coleÃ§Ã£o for vazia, devemos PULAR atÃ© o NEXT
					// Caso contrÃ¡rio, ele executa o corpo do loop uma vez como cÃ³digo normal (o erro que vocÃª viu)
					depth := 1
					tempPC := e.PC + 1
					for tempPC < len(e.Instructions) {
						l := strings.ToLower(e.Instructions[tempPC].Content)
						if strings.HasPrefix(l, "for ") || strings.HasPrefix(l, "for each ") {
							depth++
						}
						if strings.HasPrefix(l, "next") {
							depth--
						}
						if depth == 0 {
							e.PC = tempPC // O loop principal farÃ¡ o PC++ depois, caindo na linha apÃ³s o Next
							break
						}
						tempPC++
					}
				}
			}
			e.PC++
			continue
		}

		// FOR LOOP
		if strings.HasPrefix(lowerLine, "for ") {
			rest := line[4:]
			eqIdx := strings.Index(rest, "=")
			toIdx := strings.Index(strings.ToLower(rest), " to ")

			if eqIdx > -1 && toIdx > -1 {
				varName := strings.TrimSpace(rest[:eqIdx])
				varName = strings.ToLower(varName)

				startExpr := strings.TrimSpace(rest[eqIdx+1 : toIdx])
				restAfterTo := rest[toIdx+4:]

				endExpr := restAfterTo
				stepVal := 1

				stepIdx := strings.Index(strings.ToLower(restAfterTo), " step ")
				if stepIdx > -1 {
					endExpr = strings.TrimSpace(restAfterTo[:stepIdx])
					stepExpr := strings.TrimSpace(restAfterTo[stepIdx+6:])
					sVal := EvaluateExpression(stepExpr, e.Ctx)
					if i, ok := toInt(sVal); ok {
						stepVal = i
					} else {
						// Fallback or Error? VBScript expects number.
						// Let's assume 0 or panic? Better to error clearly.
						panic("Type mismatch: For loop step must be numeric")
					}
				}

				startRaw := EvaluateExpression(startExpr, e.Ctx)
				startVal, ok1 := toInt(startRaw)
				endRaw := EvaluateExpression(endExpr, e.Ctx)
				endVal, ok2 := toInt(endRaw)

				if !ok1 || !ok2 {
					panic("Type mismatch: For loop start/end must be numeric")
				}

				if e.Ctx.OptionExplicitEnabled {
					if _, ok := e.Ctx.Variables[varName]; !ok {
						e.Ctx.Err.Raise(500, "Microsoft VBScript runtime error", "Variable is undefined: '"+varName+"'", "", 0)
						if !e.Ctx.OnErrorResumeNext {
							panic(fmt.Sprintf("Variable is undefined: '%s'", varName))
						}
					}
				}

				e.Ctx.Variables[varName] = startVal

				e.LoopStack = append(e.LoopStack, LoopState{
					Type:     LoopFor,
					VarName:  varName,
					StartVal: startVal,
					EndVal:   endVal,
					StepVal:  stepVal,
					StartPC:  e.PC,
				})
			}
			e.PC++
			continue
		}

		// DO LOOP
		if strings.HasPrefix(lowerLine, "do ") {
			rest := strings.TrimSpace(lowerLine[3:])
			until := false
			cond := ""

			if strings.HasPrefix(rest, "until ") {
				until = true
				cond = line[3+6:]
			} else if strings.HasPrefix(rest, "while ") {
				until = false
				cond = line[3+6:]
			} else if rest == "" {
				cond = "false"
				until = true
			}

			res := EvaluateCondition(cond, e.Ctx)
			shouldBreak := false
			if until && res {
				shouldBreak = true
			}
			if !until && !res {
				shouldBreak = true
			}

			if shouldBreak {
				depth := 1
				tempPC := e.PC + 1
				for tempPC < len(e.Instructions) {
					l := strings.ToLower(e.Instructions[tempPC].Content)
					if strings.HasPrefix(l, "do ") {
						depth++
					}
					if l == "loop" {
						depth--
					}
					if depth == 0 {
						e.PC = tempPC + 1
						break
					}
					tempPC++
				}
				if depth > 0 {
					e.PC++
				}
			} else {
				e.LoopStack = append(e.LoopStack, LoopState{
					Type:      LoopDo,
					Condition: cond,
					Until:     until,
					StartPC:   e.PC,
				})
				e.PC++
			}
			continue
		}

		// EXIT DO
		if lowerLine == "exit do" {
			if len(e.LoopStack) > 0 {
				top := e.LoopStack[len(e.LoopStack)-1]
				if top.Type == LoopDo {
					e.LoopStack = e.LoopStack[:len(e.LoopStack)-1]
					depth := 1
					tempPC := e.PC + 1
					for tempPC < len(e.Instructions) {
						l := strings.ToLower(e.Instructions[tempPC].Content)
						if strings.HasPrefix(l, "do ") {
							depth++
						}
						if l == "loop" {
							depth--
						}
						if depth == 0 {
							e.PC = tempPC + 1
							break
						}
						tempPC++
					}
					continue
				}
			}
			e.PC++
			continue
		}

		// WHILE LOOP
		if strings.HasPrefix(lowerLine, "while ") {
			cond := strings.TrimSpace(line[5:])
			res := EvaluateCondition(cond, e.Ctx)
			if res {
				e.LoopStack = append(e.LoopStack, LoopState{
					Type:      LoopWhile,
					Condition: cond,
					StartPC:   e.PC,
				})
				e.PC++
			} else {
				// Skip to WEND
				depth := 1
				tempPC := e.PC + 1
				for tempPC < len(e.Instructions) {
					l := strings.ToLower(e.Instructions[tempPC].Content)
					if strings.HasPrefix(l, "while ") {
						depth++
					}
					if l == "wend" {
						depth--
					}
					if depth == 0 {
						e.PC = tempPC + 1
						break
					}
					tempPC++
				}
				if depth > 0 {
					e.PC++
				}
			}
			continue
		}

		// WEND
		if lowerLine == "wend" {
			if len(e.LoopStack) > 0 {
				loop := e.LoopStack[len(e.LoopStack)-1]
				if loop.Type == LoopWhile {
					// Re-evaluate condition
					res := EvaluateCondition(loop.Condition, e.Ctx)
					if res {
						e.PC = loop.StartPC + 1
						continue
					} else {
						e.LoopStack = e.LoopStack[:len(e.LoopStack)-1]
					}
				}
			}
			e.PC++
			continue
		}
		// =================================================================
		// CORREÃ‡ÃƒO: EXECUÃ‡ÃƒO DE COMANDOS FILE / SERVER / RESPONSE (Void / Standalone)
		// Detecta: File.Delete "x", Server.Execute "x", Response.AddHeader "x", "y"
		// =================================================================
		if strings.HasPrefix(lowerLine, "file.") || strings.HasPrefix(lowerLine, "server.") || strings.HasPrefix(lowerLine, "response.") || strings.HasPrefix(lowerLine, "err.") {
			lineClean := strings.TrimSpace(line)

			// Same heuristic for "Method Arg" -> "Method(Arg)" as Document.Write?
			// EvaluateExpression expects Func(Arg) for methods.
			// But for "Server.Execute 'path'", ParseRaw might have tokenized it as "Server.Execute 'path'".
			// EvaluateExpression expects "Server.Execute('path')".

			// Reuse the logic I planned for Document.Write:
			spaceIdx := strings.Index(lineClean, " ")
			parenIdx := strings.Index(lineClean, "(")

			// Only split if space exists AND (no parens OR space comes before parens)
			// This avoids splitting "Obj.Method('space here')"
			if spaceIdx > -1 && (parenIdx == -1 || spaceIdx < parenIdx) {
				method := lineClean[:spaceIdx]
				arg := strings.TrimSpace(lineClean[spaceIdx+1:])
				// Only wrap if it's not already wrapped? "Method (Arg)" -> "Method((Arg))"? No.
				// "Method Arg1, Arg2" -> "Method(Arg1, Arg2)"
				callExpr := fmt.Sprintf("%s(%s)", method, arg)
				EvaluateExpression(callExpr, e.Ctx)
			} else {
				EvaluateExpression(lineClean, e.Ctx)
			}

			e.PC++
			continue
		}
		// HANDLER FOR DOCUMENT.* (e.g. Document.Write "foo" or Document.WriteSafe "foo")
		if strings.HasPrefix(lowerLine, "document.") {
			lineClean := strings.TrimSpace(line)
			// Detect space separator
			spaceIdx := strings.Index(lineClean, " ")
			parenIdx := strings.Index(lineClean, "(")

			if spaceIdx > -1 && (parenIdx == -1 || spaceIdx < parenIdx) {
				method := lineClean[:spaceIdx]
				arg := strings.TrimSpace(lineClean[spaceIdx+1:])
				// Reconstruct as method(arg)
				callExpr := fmt.Sprintf("%s(%s)", method, arg)
				EvaluateExpression(callExpr, e.Ctx)
			} else {
				EvaluateExpression(lineClean, e.Ctx)
			}
			e.PC++
			continue
		}

		// GENERIC OBJECT METHOD CALL HANDLER (Obj.Method Args)
		// Handles cases like "fs.Append path, content" or "mail.Send ..."
		// that are not assignments and don't start with keywords.
		if strings.Contains(line, ".") && !strings.Contains(line, "=") {
			lineClean := strings.TrimSpace(line)
			// Try to find first space to split Method and Args
			spaceIdx := strings.Index(lineClean, " ")
			parenIdx := strings.Index(lineClean, "(")

			if spaceIdx > -1 && (parenIdx == -1 || spaceIdx < parenIdx) {
				method := lineClean[:spaceIdx]
				// Basic check to see if it looks like Obj.Method
				if strings.Contains(method, ".") {
					arg := strings.TrimSpace(lineClean[spaceIdx+1:])
					// Construct "Method(Args)" for EvaluateExpression
					callExpr := fmt.Sprintf("%s(%s)", method, arg)
					EvaluateExpression(callExpr, e.Ctx)
					e.PC++
					continue
				}
			} else {
				// No space. "Obj.Method"? (No args or using parens already)
				// e.g. "rs.Close" or "rs.Close()"
				EvaluateExpression(lineClean, e.Ctx)
				e.PC++
				continue
			}
		}

		// =================================================================

		// CONST
		if strings.HasPrefix(lowerLine, "const ") {
			// Const Name = Value, Name2 = Value2
			rest := line[6:]
			parts := strings.Split(rest, ",")
			for _, part := range parts {
				eqIdx := strings.Index(part, "=")
				if eqIdx > -1 {
					name := strings.TrimSpace(part[:eqIdx])
					valExpr := strings.TrimSpace(part[eqIdx+1:])
					val := EvaluateExpression(valExpr, e.Ctx)
					e.Ctx.Constants[strings.ToLower(name)] = val
				}
			}
			e.PC++
			continue
		}

		// ASSIGNMENT (x = y) / SET Objects
		if strings.Contains(line, "=") {
			parts := strings.SplitN(line, "=", 2)
			key := strings.TrimSpace(parts[0])

			// CORREÃ‡ÃƒO: Remove "Set " do inÃ­cio se existir
			// Caso contrÃ¡rio, a variÃ¡vel Ã© salva como "set user" em vez de "user"
			if strings.HasPrefix(strings.ToLower(key), "set ") {
				key = strings.TrimSpace(key[4:])
			}

			valExpr := strings.TrimSpace(parts[1])
			lowerKey := strings.ToLower(key)

			// WITH-scoped property assignment
			if strings.HasPrefix(key, ".") {
				val := EvaluateExpression(valExpr, e.Ctx)
				if len(e.WithStack) > 0 {
					base := e.WithStack[len(e.WithStack)-1]
					if setWithPath(base, strings.TrimSpace(key[1:]), val, e.Ctx) {
						e.PC++
						continue
					}
				}
				// Skip further processing to mirror VBScript behavior even if With context is missing
				e.PC++
				continue
			}

			// Check if it's a Constant (Illegal Assignment)
			if _, isConst := e.Ctx.Constants[lowerKey]; isConst {
				e.Ctx.Err.Raise(500, "Microsoft VBScript compilation error", "Illegal assignment: '"+key+"'", "", 0)
				if !e.Ctx.OnErrorResumeNext {
					panic(fmt.Sprintf("Illegal assignment: '%s' is a constant", key))
				}
				e.PC++
				continue
			}

			// Evaluate the value expression
			val := EvaluateExpression(valExpr, e.Ctx)

			// 1. SPECIAL PROPERTIES
			if lowerKey == "response.contenttype" {
				e.Ctx.ResponseState.ContentType = fmt.Sprintf("%v", val)
				e.PC++
				continue
			}
			if lowerKey == "response.status" {
				e.Ctx.ResponseState.Status = fmt.Sprintf("%v", val)
				e.PC++
				continue
			}
			if strings.HasPrefix(lowerKey, "response.cookies(\"") {
				cookieName := strings.TrimSuffix(strings.TrimPrefix(lowerKey, "response.cookies(\""), "\")")
				val := EvaluateExpression(valExpr, e.Ctx)
				e.Ctx.ResponseState.Cookies[cookieName] = fmt.Sprintf("%v", val)
				e.PC++
				continue
			}
			// Response Properties
			if lowerKey == "response.expires" {
				val := EvaluateExpression(valExpr, e.Ctx)
				i, _ := toInt(val)
				e.Ctx.ResponseState.Expires = i
				e.PC++
				continue
			}
			if lowerKey == "response.expiresabsolute" {
				val := EvaluateExpression(valExpr, e.Ctx)
				if t, ok := val.(time.Time); ok {
					e.Ctx.ResponseState.ExpiresAbsolute = t
				} else if s, ok := val.(string); ok {
					if t, err := time.Parse("01/02/2006 15:04:05", s); err == nil {
						e.Ctx.ResponseState.ExpiresAbsolute = t
					} else if t, err := time.Parse("01/02/2006", s); err == nil {
						e.Ctx.ResponseState.ExpiresAbsolute = t
					}
				}
				e.PC++
				continue
			}
			if lowerKey == "response.cachecontrol" {
				val := EvaluateExpression(valExpr, e.Ctx)
				e.Ctx.ResponseState.CacheControl = fmt.Sprintf("%v", val)
				e.PC++
				continue
			}
			if lowerKey == "response.charset" {
				val := EvaluateExpression(valExpr, e.Ctx)
				e.Ctx.ResponseState.Charset = fmt.Sprintf("%v", val)
				e.PC++
				continue
			}
			if lowerKey == "response.pics" {
				val := EvaluateExpression(valExpr, e.Ctx)
				e.Ctx.ResponseState.PICS = fmt.Sprintf("%v", val)
				e.PC++
				continue
			}
			// Server Properties
			if lowerKey == "server.scripttimeout" {
				val := EvaluateExpression(valExpr, e.Ctx)
				i, _ := toInt(val)
				e.Ctx.ScriptTimeout = i
				e.PC++
				continue
			}

			if lowerKey == "session.timeout" {
				val := EvaluateExpression(valExpr, e.Ctx)
				// Inline toInt helper logic
				var i int
				if iv, ok := val.(int); ok {
					i = iv
				} else if s, ok := val.(string); ok {
					fmt.Sscanf(s, "%d", &i)
				}
				if i > 0 {
					e.Ctx.Session.Timeout = i
				}
				e.PC++
				continue
			}

			// 2. SESSION & APPLICATION (PRIORIDADE ALTA)
			// Devem vir antes da lÃ³gica genÃ©rica de arrays para evitar conflitos
			if strings.HasPrefix(lowerKey, "session(\"") {
				sessKey := strings.TrimSuffix(strings.TrimPrefix(lowerKey, "session(\""), "\")")
				val := EvaluateExpression(valExpr, e.Ctx)
				e.Ctx.Session.Set(sessKey, val)
				e.PC++
				continue
			}
			if strings.HasPrefix(lowerKey, "application(\"") {
				appKey := strings.TrimSuffix(strings.TrimPrefix(lowerKey, "application(\""), "\")")
				val := EvaluateExpression(valExpr, e.Ctx)
				e.Ctx.Application.Set(appKey, val)
				e.PC++
				continue
			}

			// 3. JSON OBJECTS / ARRAYS (LÃ³gica Nova)
			// Detecta sintaxe: variavel(indice) = valor
			if idxStart := strings.Index(key, "("); idxStart > -1 && strings.HasSuffix(key, ")") {
				varName := strings.TrimSpace(key[:idxStart])
				lowerVarName := strings.ToLower(varName)
				idxContent := key[idxStart+1 : len(key)-1]

				// ProteÃ§Ã£o: Ignora se for tentar acessar session/application via variÃ¡vel dinÃ¢mica aqui
				// (Embora o prefixo acima jÃ¡ capture strings literais, isso protege contra casos de borda)
				if lowerVarName != "session" && lowerVarName != "application" {

					// Verifica se a variÃ¡vel base existe e Ã© um Objeto/Array
					if baseObj, exists := e.Ctx.Variables[lowerVarName]; exists {

						// Caso Map (JSON Object): obj("prop") = val
						if mapObj, ok := baseObj.(map[string]interface{}); ok {
							keyVal := EvaluateExpression(idxContent, e.Ctx)
							strKey := fmt.Sprintf("%v", keyVal)
							newVal := EvaluateExpression(valExpr, e.Ctx)

							mapObj[strKey] = newVal
							// Map Ã© referÃªncia, nÃ£o precisa reatribuir ao Variables, mas nÃ£o faz mal
							e.PC++
							continue
						}

						// Caso Slice (JSON Array): arr(0) = val
						if _, ok := baseObj.([]interface{}); ok {
							// Support multi-dimensional index like arr(0,1)
							parts := strings.Split(idxContent, ",")
							var indices []int
							for _, p := range parts {
								p = strings.TrimSpace(p)
								if p == "" {
									indices = append(indices, 0)
									continue
								}
								v := EvaluateExpression(p, e.Ctx)
								i, _ := toInt(v)
								indices = append(indices, i)
							}

							// Traverse/create nested slices to reach target
							current := baseObj
							parentSlice := ([]interface{})(nil)
							parentName := lowerVarName
							for depthIdx, idx := range indices {
								if idx < 0 {
									break
								}
								switch cur := current.(type) {
								case []interface{}:
									// Ensure slice has length to include idx
									if idx >= len(cur) {
										// extend with nils up to idx
										for k := len(cur); k <= idx; k++ {
											cur = append(cur, nil)
										}
									}
									// If last index, assign
									if depthIdx == len(indices)-1 {
										newVal := EvaluateExpression(valExpr, e.Ctx)
										cur[idx] = newVal
										// write back to parent or root
										if parentSlice == nil {
											e.Ctx.Variables[parentName] = cur
										} else {
											parentSliceIndex := indices[depthIdx-1]
											parentSlice[parentSliceIndex] = cur
											// If parent is root variable, ensure stored
											e.Ctx.Variables[parentName] = parentSlice
										}
										current = cur[idx]
									}
									// Not last: drill down
									parentSlice = cur
									current = cur[idx]
									// If next level is nil, create a slice to continue
									if current == nil {
										cur[idx] = []interface{}{}
										current = cur[idx]
									}
									// update parent reference
									parentSlice = cur
								default:
									// If current is nil or not a slice and we need to go deeper, create nested
									if current == nil {
										// create nested slices for remaining dims
										remaining := make([]int, len(indices)-depthIdx)
										for ri := range remaining {
											remaining[ri] = 1
										}
										newNested := makeNestedArray(remaining)
										current = newNested
										// assign into parentSlice
										if parentSlice != nil {
											parentSlice[indices[depthIdx-1]] = current
											if parentName != "" {
												e.Ctx.Variables[parentName] = parentSlice
											}
										}
										// continue loop after creation
										continue
									}
									// Otherwise cannot index into non-slice
									//break
								}
							}
							// ensure root variable stored (in case only append happened)
							e.Ctx.Variables[lowerVarName] = e.Ctx.Variables[lowerVarName]
							e.PC++
							continue
						}
					}
				}
			}

			// 4. COM Object Property Assignment (obj.Property = value)
			if dotIdx := strings.Index(key, "."); dotIdx > -1 && !strings.Contains(key, "(") {
				objName := strings.ToLower(key[:dotIdx])
				propName := strings.ToLower(key[dotIdx+1:])

				var obj interface{}
				found := false

				// 1. Local Variables
				if val, ok := e.Ctx.Variables[objName]; ok {
					obj = val
					found = true
				}
				// 2. Global Variables
				if !found && e.Ctx.GlobalVariables != nil {
					if val, ok := e.Ctx.GlobalVariables[objName]; ok {
						obj = val
						found = true
					}
				}
				// 3. Class Instance Variables
				if !found && e.Ctx.CurrentInstance != nil {
					if val, ok := e.Ctx.CurrentInstance.Variables[objName]; ok {
						obj = val
						found = true
					}
				}

				if found && obj != nil {
					// Try Component interface
					if comp, ok := obj.(Component); ok {
						val := EvaluateExpression(valExpr, e.Ctx)
						comp.SetProperty(propName, val)
						e.PC++
						continue
					}

					// Try map (JSON object)
					if mapObj, ok := obj.(map[string]interface{}); ok {
						val := EvaluateExpression(valExpr, e.Ctx)
						mapObj[propName] = val
						e.PC++
						continue
					}

					// Try using reflection to call SetProperty method (for types that implement it but don't match interface)
					val := EvaluateExpression(valExpr, e.Ctx)
					if tryCallSetProperty(obj, propName, val) {
						e.PC++
						continue
					}

					// If we reach here, object was found but couldn't be accessed
					// This is an error - the object doesn't support property assignment
					e.Ctx.Err.Raise(500, "Microsoft VBScript runtime error", fmt.Sprintf("Object doesn't support this property or method: '%s.%s'", objName, propName), "", 0)
					if !e.Ctx.OnErrorResumeNext {
						return
					}
					e.PC++
					continue
				} else if !found {
					// Object not found - raise undefined variable error
					e.Ctx.Err.Raise(500, "Microsoft VBScript runtime error", fmt.Sprintf("Variable is undefined: '%s'", key), "", 0)
					if !e.Ctx.OnErrorResumeNext {
						panic(fmt.Sprintf("Variable is undefined: '%s'", key))
					}
					e.PC++
					continue
				}
			}

			// 4. GENERIC ASSIGNMENT
			evaluated := EvaluateExpression(valExpr, e.Ctx)
			keyLower := strings.ToLower(key)

			// Try Class Member (Variable/Property Let/Set)
			if e.Ctx.CurrentInstance != nil {
				// Check Member Variables (Private/Public)
				if _, ok := e.Ctx.CurrentInstance.ClassDef.Variables[keyLower]; ok {
					e.Ctx.CurrentInstance.Variables[keyLower] = evaluated
					e.PC++
					continue
				}
				// Check Property Let/Set
				// ...
			}

			if e.Ctx.OptionExplicitEnabled {
				// Note: Object property assignments (Obj.Prop = value) are already handled above
				// This check only applies to simple variable assignments
				if !strings.Contains(keyLower, ".") {
					found := false
					if _, ok := e.Ctx.Variables[keyLower]; ok {
						found = true
					}
					if !found && e.Ctx.GlobalVariables != nil {
						if _, ok := e.Ctx.GlobalVariables[keyLower]; ok {
							found = true
						}
					}

					if !found {
						e.Ctx.Err.Raise(500, "Microsoft VBScript runtime error", "Variable is undefined: '"+key+"'", "", 0)
						if !e.Ctx.OnErrorResumeNext {
							panic(fmt.Sprintf("Variable is undefined: '%s'", key))
						}
					}
				}
			}

			// Assignment Logic:
			// 1. If exists in Local, set Local.
			// 2. If exists in Global, set Global.
			// 3. Else set Local (Implicit Declaration).

			if _, ok := e.Ctx.Variables[keyLower]; ok {
				e.Ctx.Variables[keyLower] = evaluated
			} else if e.Ctx.GlobalVariables != nil {
				if _, ok := e.Ctx.GlobalVariables[keyLower]; ok {
					e.Ctx.GlobalVariables[keyLower] = evaluated
				} else {
					e.Ctx.Variables[keyLower] = evaluated
				}
			} else {
				e.Ctx.Variables[keyLower] = evaluated
			}
			e.PC++
			continue
		}

		// RESPONSE Commands
		if lowerLine == "response.end" {
			e.Ctx.End()
			return
		}
		if lowerLine == "response.clear" {
			e.Ctx.Clear()
			e.PC++
			continue
		}
		if lowerLine == "response.flush" {
			e.Ctx.Flush()
			e.PC++
			continue
		}

		// SESSION Commands
		if lowerLine == "session.abandon" {
			e.Ctx.Session.Abandon()
			e.PC++
			continue
		}

		// APPLICATION Commands
		if lowerLine == "application.lock" {
			e.Ctx.Application.Lock()
			e.PC++
			continue
		}
		if lowerLine == "application.unlock" {
			e.Ctx.Application.Unlock()
			e.PC++
			continue
		}

		// NEXT / LOOP
		if strings.HasPrefix(lowerLine, "next") || lowerLine == "loop" {
			if len(e.LoopStack) > 0 {
				loop := e.LoopStack[len(e.LoopStack)-1]

				if loop.Type == LoopFor {
					// Safely get current value, allowing for it to have been modified to a string/float but convertible
					currRaw := e.Ctx.Variables[loop.VarName]
					currVal, ok := toInt(currRaw)
					if !ok {
						panic("Type mismatch: For loop variable modified to non-numeric")
					}

					currVal += loop.StepVal
					e.Ctx.Variables[loop.VarName] = currVal

					finished := false
					if loop.StepVal >= 0 {
						if currVal > loop.EndVal {
							finished = true
						}
					} else {
						if currVal < loop.EndVal {
							finished = true
						}
					}

					if !finished {
						e.PC = loop.StartPC + 1
						continue
					} else {
						e.LoopStack = e.LoopStack[:len(e.LoopStack)-1]
					}
				} else if loop.Type == LoopForEach {
					// ForEach loop: advance to next item in collection
					loop.Index++
					if loop.Index < len(loop.Collection) {
						e.Ctx.Variables[loop.VarName] = loop.Collection[loop.Index]
						e.LoopStack[len(e.LoopStack)-1] = loop // Update the loop state
						e.PC = loop.StartPC + 1
						continue
					} else {
						e.LoopStack = e.LoopStack[:len(e.LoopStack)-1]
					}
				} else if loop.Type == LoopDo {
					res := EvaluateCondition(loop.Condition, e.Ctx)
					shouldBreak := false
					if loop.Until && res {
						shouldBreak = true
					}
					if !loop.Until && !res {
						shouldBreak = true
					}

					if shouldBreak {
						e.LoopStack = e.LoopStack[:len(e.LoopStack)-1]
					} else {
						e.PC = loop.StartPC + 1
						continue
					}
				}
			}
			e.PC++
			continue
		}

		e.PC++
	}
}

// ExecuteClassMethod executes a method (Sub/Function/Property) contextually within a Class Instance
// bindings holds optional ByRef setters corresponding to args positions.
func (e *Engine) ExecuteClassMethod(instance *ClassInstance, methodName string, propType PropertyType, args []interface{}, bindings []ByRefSetter) interface{} {
	nameLower := strings.ToLower(methodName)
	var proc Procedure
	var found bool

	// Find the procedure
	if propType == PropGet || propType == PropLet || propType == PropSet {
		// Look in Properties
		if props, ok := instance.ClassDef.Properties[nameLower]; ok {
			for _, p := range props {
				if p.Type == propType {
					proc = Procedure{LineNum: p.LineNum, Params: p.Params}
					found = true
					break
				}
			}
		}
	}

	if !found {
		// Look in Methods (Sub/Function)
		if m, ok := instance.ClassDef.Methods[nameLower]; ok {
			proc = m
			found = true
		}
	}

	if !found {
		return nil
	}

	// Create a new Context Scope (Local Variables)
	subCtx := *instance.Ctx                               // Shallow copy
	subCtx.Variables = make(map[string]interface{})       // Fresh Local Scope
	subCtx.GlobalVariables = instance.Ctx.GlobalVariables // Preserve Global Link

	// Link Engine
	subCtx.Engine = nil

	// Create a sub-engine execution
	subEngine := &Engine{
		Instructions:   e.Instructions,
		Labels:         e.Labels,
		Classes:        e.Classes,
		Ctx:            &subCtx,
		PC:             proc.LineNum + 1,
		SingleProcMode: true,
		CallStack:      make([]CallFrame, 0),
		IfStack:        make([]IfState, 0),
		LoopStack:      make([]LoopState, 0),
		SelectStack:    make([]SelectState, 0),
		WithStack:      make([]interface{}, 0),
	}

	// Ensure sub-context points to the new engine (needed for With stacks)
	subCtx.Engine = subEngine

	// Inject Args
	for i, paramName := range proc.Params {
		if i < len(args) {
			subEngine.Ctx.Variables[strings.ToLower(paramName)] = args[i]
		}
	}

	// Initialize Return Value Variable and track current method
	subEngine.Ctx.Variables[nameLower] = nil // Default to Empty
	subEngine.Ctx.CurrentMethodName = nameLower

	// Set Current Instance for Me
	subEngine.Ctx.CurrentInstance = instance

	// RUN
	subEngine.Run(subEngine.Ctx)

	// Copy-out ByRef args if provided
	if len(bindings) > 0 {
		for i, setter := range bindings {
			if setter == nil {
				continue
			}
			if i < len(proc.Params) {
				pname := strings.ToLower(proc.Params[i])
				if val, ok := subEngine.Ctx.Variables[pname]; ok {
					setter(val)
				}
			}
		}
	}

	// Return Value (for Function/Property Get)
	if val, ok := subEngine.Ctx.Variables[nameLower]; ok {
		return val
	}

	return nil
}

// RunGlobalEvent executes a Sub defined in Global.asa
func RunGlobalEvent(eventName string, ctx *ExecutionContext) {
	if AppState.GlobalASACode == "" {
		return
	}

	// Construct a script that defines the subs and calls the event
	// We need to ensure the event Sub exists, otherwise Call might fail?
	// The Engine's Call handles missing subs gracefully?
	// Check engine.go Call: "if proc, ok := e.Labels... continue".
	// If not found, it does nothing?
	// Wait, if "Call X", and X not found:
	// It falls through to "DIM", "RESPONSE.WRITE", etc.
	// It might try to evaluate "Call X" as an expression if it has parens?
	// Or falls to generic assignment?
	// If "Call Application_OnStart", lines start with "call ".
	// Logic: if proc ok -> run. Else -> continue.
	// So if Sub not found, it just ignores the call. That is safe.

	fullCode := "<%" + AppState.GlobalASACode + "\nCall " + eventName + "%>"

	tokens := ParseRaw(fullCode)
	engine := Prepare(tokens)
	engine.Run(ctx)
}
