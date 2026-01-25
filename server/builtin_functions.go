package server

import (
	"fmt"
	"math"
	"math/rand"
	"os"
	"reflect"
	"strconv"
	"strings"
	"time"

	vb "github.com/guimaraeslucas/vbscript-go"
	"github.com/guimaraeslucas/vbscript-go/ast"
)

// EmptyValue represents VBScript Empty type
type EmptyValue struct{}

// NothingValue represents VBScript Nothing type (null object reference)
type NothingValue struct{}

// NullValue represents VBScript Null type (no valid data)
type NullValue struct{}

// evalBuiltInFunction evaluates VBScript built-in functions
// Returns (result, wasHandled) where wasHandled indicates if the function was recognized
func evalBuiltInFunction(funcName string, args []interface{}, ctx *ExecutionContext) (interface{}, bool) {
	funcLower := strings.ToLower(funcName)

	// Try date/time functions first
	if result, handled := evalDateTimeFunction(funcLower, args, ctx); handled {
		return result, true
	}

	switch funcLower {
	// Dynamic code execution
	case "executeglobal":
		// ExecuteGlobal(code) - executes VBScript code in the current global scope
		if len(args) == 0 {
			return nil, true
		}
		code := toString(args[0])
		if code == "" {
			return nil, true
		}

		// Parse the code
		parser := vb.NewParser(code)
		var program *ast.Program
		func() {
			defer func() {
				if r := recover(); r != nil {
					// Parse error - silently fail as per VBScript behavior
					program = nil
				}
			}()
			program = parser.Parse()
		}()

		if program == nil {
			return nil, true
		}

		// Execute in the current context (global scope)
		visitor := NewASPVisitor(ctx, nil)
		for _, stmt := range program.Body {
			if err := visitor.VisitStatement(stmt); err != nil {
				// ExecuteGlobal errors are typically swallowed in VBScript
				// but we can log them for debugging
				return nil, true
			}
		}
		return nil, true

	case "execute":
		// Execute(code) - executes VBScript code in a local scope
		// For now, we'll treat it the same as ExecuteGlobal since we don't have proper scope isolation
		if len(args) == 0 {
			return nil, true
		}
		code := toString(args[0])
		if code == "" {
			return nil, true
		}

		// Parse the code
		parser := vb.NewParser(code)
		var program *ast.Program
		func() {
			defer func() {
				if r := recover(); r != nil {
					program = nil
				}
			}()
			program = parser.Parse()
		}()

		if program == nil {
			return nil, true
		}

		// Execute in the current context
		visitor := NewASPVisitor(ctx, nil)
		for _, stmt := range program.Body {
			if err := visitor.VisitStatement(stmt); err != nil {
				return nil, true
			}
		}
		return nil, true

	// Type checking functions
	case "isempty":
		if len(args) == 0 {
			return true, true
		}
		val := args[0]
		// Check if value is Empty (nil or EmptyValue)
		// NOTE: nil represents an uninitialized variable (Empty), NOT Null
		if val == nil {
			return true, true
		}
		if _, ok := val.(EmptyValue); ok {
			return true, true
		}
		// NullValue is NOT empty
		if _, ok := val.(NullValue); ok {
			return false, true
		}
		// Check if it's a newly declared variable (stored as nil in context)
		return false, true

	case "isnull":
		if len(args) == 0 {
			return false, true
		}
		val := args[0]
		// In VBScript, Null is a special value indicating no valid data
		// We represent it as NullValue{}
		if _, ok := val.(NullValue); ok {
			return true, true
		}
		// nil is NOT Null - it's Empty
		return false, true

	case "isnothing":
		if len(args) == 0 {
			return false, true
		}
		val := args[0]
		// Check if value is Nothing (null object reference)
		if val == nil {
			return true, true
		}
		if _, ok := val.(NothingValue); ok {
			return true, true
		}
		return false, true

	case "isobject":
		if len(args) == 0 {
			return false, true
		}
		val := args[0]
		// Check if value is an object (map, array, component, library)
		switch val.(type) {
		case map[string]interface{}, []interface{}, ASPLibrary, *VBArray:
			return true, true
		default:
			if val != nil {
				rv := reflect.ValueOf(val)
				if rv.Kind() == reflect.Ptr {
					rv = rv.Elem()
				}
				return rv.Kind() == reflect.Struct || rv.Kind() == reflect.Map, true
			}
			return false, true
		}

	case "typename":
		if len(args) == 0 {
			return "Empty", true
		}
		val := args[0]
		return getTypeName(val), true

	case "vartype":
		if len(args) == 0 {
			return 0, true // vbEmpty
		}
		val := args[0]
		return getVarType(val), true

	case "scriptengine":
		// ScriptEngine() - Returns "VBScript" for compatibility
		return "VBScript", true

	case "scriptenginebuildversion":
		// ScriptEngineBuildVersion() - Returns VBScript build version (18702 = VBScript 5.8)
		return 18702, true

	case "scriptenginemajorversion":
		// ScriptEngineMajorVersion() - Returns 5 (VBScript 5.x)
		return 5, true

	case "scriptengineminorversion":
		// ScriptEngineMinorVersion() - Returns 8 (VBScript 5.8)
		return 8, true

	case "ascw":
		if len(args) == 0 {
			return 0, true
		}
		s := toString(args[0])
		runes := []rune(s)
		if len(runes) == 0 {
			return 0, true
		}
		return int(runes[0]), true

	case "chrw":
		if len(args) == 0 {
			return "", true
		}
		code := toInt(args[0])
		return string(rune(code)), true

	case "env":
		// Env(name) - Returns environment variable value
		if len(args) == 0 {
			return "", true
		}
		envName := toString(args[0])
		// Try to get from OS environment first
		value := os.Getenv(envName)
		// Return the value (empty string if not found, which matches VBScript behavior)
		return value, true

	case "eval":
		// Eval(expression) - Evaluate expression string and return result
		if len(args) == 0 {
			return nil, true
		}
		exprStr := toString(args[0])
		// Evaluate the expression string using the context's evaluation mechanism
		// This requires access to the context's expression evaluator
		// For now, return the parsed result
		result := evalExpression(exprStr, ctx)
		return result, true

	case "getobject":
		if ctx == nil || ctx.Server == nil {
			return nil, true
		}
		if len(args) == 0 {
			return nil, true
		}

		pathArg := toString(args[0])
		className := ""
		if len(args) > 1 {
			className = toString(args[1])
		}

		progID := strings.TrimSpace(className)
		if progID == "" {
			progID = strings.TrimSpace(pathArg)
		}
		if progID == "" {
			return nil, true
		}

		obj, err := ctx.Server.CreateObject(progID)
		if err != nil {
			ctx.Err.SetError(err)
			return nil, true
		}

		if pathArg != "" && className != "" {
			if loader, ok := obj.(interface{ Load(string) error }); ok {
				if err := loader.Load(pathArg); err != nil {
					ctx.Err.SetError(err)
				}
			} else if opener, ok := obj.(interface{ Open(string) error }); ok {
				if err := opener.Open(pathArg); err != nil {
					ctx.Err.SetError(err)
				}
			}
		}

		return obj, true

	case "rgb":
		// RGB(red, green, blue) - returns color as integer
		if len(args) < 3 {
			return 0, true
		}
		r := toInt(args[0]) % 256
		g := toInt(args[1]) % 256
		b := toInt(args[2]) % 256
		if r < 0 {
			r = 0
		}
		if g < 0 {
			g = 0
		}
		if b < 0 {
			b = 0
		}
		// VBScript RGB returns as BGR integer (Blue in low byte)
		return r + (g << 8) + (b << 16), true
	// Array functions
	case "array":
		// Array(elem1, elem2, ...) - creates an array
		result := make([]interface{}, len(args))
		for i, arg := range args {
			result[i] = arg
		}
		return NewVBArrayFromValues(0, result), true

	case "isarray":
		// IsArray(var) - checks if variable is an array
		if len(args) == 0 {
			return false, true
		}
		val := args[0]
		_, ok := toVBArray(val)
		return ok, true

	case "lbound":
		// LBound(array, [dimension]) - returns lower bound respecting Option Base
		if len(args) == 0 {
			return -1, true
		}
		dim := 1
		if len(args) > 1 {
			dim = toInt(args[1])
		}
		lower, _, ok := arrayBounds(args[0], dim)
		if !ok {
			return -1, true
		}
		return lower, true

	case "ubound":
		// UBound(array, [dimension]) - returns upper bound (last index)
		if len(args) == 0 {
			return -1, true
		}
		dim := 1
		if len(args) > 1 {
			dim = toInt(args[1])
		}
		_, upper, ok := arrayBounds(args[0], dim)
		if !ok {
			return -1, true
		}
		return upper, true

	case "split":
		// Split(expression, delimiter, [limit], [compare]) - splits string into array
		if len(args) < 2 {
			return NewVBArrayFromValues(0, []interface{}{}), true
		}
		str := toString(args[0])
		delim := toString(args[1])

		// Handle limit parameter
		limit := -1
		if len(args) > 2 {
			limit = toInt(args[2])
		}

		var parts []string
		if limit > 0 {
			parts = strings.SplitN(str, delim, limit)
		} else {
			parts = strings.Split(str, delim)
		}

		result := make([]interface{}, len(parts))
		for i, part := range parts {
			result[i] = part
		}
		return NewVBArrayFromValues(0, result), true

	case "join":
		// Join(array, delimiter) - joins array elements into string
		if len(args) < 2 {
			return "", true
		}
		arr, ok := toVBArray(args[0])
		if !ok {
			return "", true
		}
		delim := toString(args[1])

		parts := make([]string, len(arr.Values))
		for i, item := range arr.Values {
			parts[i] = toString(item)
		}
		return strings.Join(parts, delim), true

	case "filter":
		// Filter(array, match, [include], [compare]) - filters string arrays
		if len(args) < 2 {
			return NewVBArrayFromValues(0, []interface{}{}), true
		}
		arr, ok := toVBArray(args[0])
		if !ok {
			return NewVBArrayFromValues(0, []interface{}{}), true
		}

		match := toString(args[1])
		include := true
		if len(args) > 2 {
			include = isTruthy(args[2])
		}

		compare := 0
		if len(args) > 3 {
			compare = toInt(args[3])
		}

		needle := match
		if compare == 1 {
			needle = strings.ToLower(match)
		}

		result := make([]interface{}, 0)
		for _, item := range arr.Values {
			itemStr := toString(item)
			text := itemStr
			if compare == 1 {
				text = strings.ToLower(itemStr)
			}

			found := needle == "" || strings.Contains(text, needle)
			if found == include {
				result = append(result, item)
			}
		}

		return NewVBArrayFromValues(0, result), true

	// String Functions
	case "len":
		// LEN(string) - returns length of string
		if len(args) == 0 {
			return 0, true
		}
		return len(toString(args[0])), true

	case "left":
		// LEFT(string, length) - returns leftmost characters
		if len(args) < 2 {
			return "", true
		}
		s := toString(args[0])
		n := toInt(args[1])
		if n > len(s) {
			n = len(s)
		}
		if n < 0 {
			n = 0
		}
		return s[:n], true

	case "right":
		// RIGHT(string, length) - returns rightmost characters
		if len(args) < 2 {
			return "", true
		}
		s := toString(args[0])
		n := toInt(args[1])
		if n > len(s) {
			n = len(s)
		}
		if n < 0 {
			n = 0
		}
		return s[len(s)-n:], true

	case "mid":
		// MID(string, start, [length]) - returns substring
		if len(args) < 2 {
			return "", true
		}
		s := toString(args[0])
		start := toInt(args[1]) - 1 // VBScript is 1-based
		length := len(s)
		if len(args) >= 3 {
			length = toInt(args[2])
		}
		if start < 0 {
			start = 0
		}
		if start >= len(s) {
			return "", true
		}
		end := start + length
		if end > len(s) {
			end = len(s)
		}
		return s[start:end], true

	case "instr":
		// INSTR([start], string1, string2) - find substring position
		var s1, s2 string
		start := 1 // VBScript default start is 1
		if len(args) == 2 {
			s1 = toString(args[0])
			s2 = toString(args[1])
		} else if len(args) >= 3 {
			start = toInt(args[0])
			s1 = toString(args[1])
			s2 = toString(args[2])
		} else {
			return 0, true
		}
		if start < 1 {
			start = 1
		}
		if start > len(s1) {
			return 0, true
		}
		s1Lower := strings.ToLower(s1)
		s2Lower := strings.ToLower(s2)
		idx := strings.Index(s1Lower[start-1:], s2Lower)
		if idx == -1 {
			// Fallback: scan full string in case of unexpected start math issues
			fullIdx := strings.Index(s1Lower, s2Lower)
			if fullIdx == -1 {
				return 0, true
			}
			return fullIdx + 1, true
		}
		return idx + start, true // Return 1-based position

	case "instrrev":
		// INSTRREV(string, substring, [start]) - find substring from right
		if len(args) < 2 {
			return 0, true
		}
		s1 := toString(args[0])
		s2 := toString(args[1])
		start := -1
		if len(args) >= 3 {
			start = toInt(args[2]) - 1 // VBScript is 1-based
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
		return idx + 1, true // Return 1-based position

	case "replace":
		// REPLACE(string, find, replace, [start], [count], [compare])
		if len(args) < 3 {
			return "", true
		}
		s := toString(args[0])
		find := toString(args[1])
		repl := toString(args[2])

		start := 1
		if len(args) >= 4 {
			start = toInt(args[3])
		}
		if start < 1 {
			start = 1
		}

		count := -1
		if len(args) >= 5 {
			count = toInt(args[4])
		}

		compare := 1 // Default to text (case-insensitive) compare
		if len(args) >= 6 {
			compare = toInt(args[5])
		}

		if find == "" || start > len(s) || count == 0 {
			return s, true
		}

		idxStart := start - 1 // convert to 0-based
		last := 0
		var b strings.Builder
		replaced := 0

		source := s
		target := find
		if compare == 1 {
			source = strings.ToLower(s)
			target = strings.ToLower(find)
		}

		for idxStart <= len(s) {
			segment := source[idxStart:]
			pos := strings.Index(segment, target)
			if pos == -1 {
				break
			}
			matchPos := idxStart + pos
			b.WriteString(s[last:matchPos])
			b.WriteString(repl)
			matchEnd := matchPos + len(find)
			last = matchEnd
			replaced++
			if count >= 0 && replaced >= count {
				break
			}
			idxStart = matchEnd
		}

		b.WriteString(s[last:])
		return b.String(), true

	case "trim":
		// TRIM(string) - removes leading and trailing spaces
		return strings.TrimSpace(toString(args[0])), true

	case "ltrim":
		// LTRIM(string) - removes leading spaces
		if len(args) == 0 {
			return "", true
		}
		return strings.TrimLeft(toString(args[0]), " \t"), true

	case "rtrim":
		// RTRIM(string) - removes trailing spaces
		if len(args) == 0 {
			return "", true
		}
		return strings.TrimRight(toString(args[0]), " \t"), true

	case "lcase":
		// LCASE(string) - converts to lowercase
		return strings.ToLower(toString(args[0])), true

	case "ucase":
		// UCASE(string) - converts to uppercase
		return strings.ToUpper(toString(args[0])), true

	case "space":
		// SPACE(number) - returns string of spaces
		n := toInt(args[0])
		if n < 0 {
			n = 0
		}
		return strings.Repeat(" ", n), true

	case "string":
		// STRING(number, character) - returns character repeated
		if len(args) < 2 {
			return "", true
		}
		num := toInt(args[0])
		if num < 0 {
			num = 0
		}

		// Accept both string and numeric character codes like VBScript
		charStr := ""
		switch v := args[1].(type) {
		case int, int8, int16, int32, int64, uint, uint8, uint16, uint32, uint64, float32, float64:
			code := toInt(v) & 0xFF
			charStr = string(rune(code))
		default:
			charStr = toString(args[1])
		}

		runes := []rune(charStr)
		if len(runes) == 0 {
			return "", true
		}

		return strings.Repeat(string(runes[0]), num), true

	case "strreverse":
		// STRREVERSE(string) - reverses string
		s := toString(args[0])
		runes := []rune(s)
		for i, j := 0, len(runes)-1; i < j; i, j = i+1, j-1 {
			runes[i], runes[j] = runes[j], runes[i]
		}
		return string(runes), true

	case "strcomp":
		// STRCOMP(string1, string2, [compare]) - compares strings
		if len(args) < 2 {
			return 0, true
		}
		s1 := toString(args[0])
		s2 := toString(args[1])
		caseInsensitive := false
		if len(args) >= 3 {
			caseInsensitive = toInt(args[2]) == 1
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

	case "asc":
		// ASC(string) - returns ASCII code of first character
		s := toString(args[0])
		if len(s) > 0 {
			return int(s[0]), true
		}
		return 0, true

	case "chr":
		// CHR(code) - returns character from ASCII code
		code := toInt(args[0])
		if code >= 0 && code <= 255 {
			return string(rune(code)), true
		}
		return "", true

	case "hex":
		// HEX(number) - converts number to hexadecimal string
		if len(args) == 0 {
			return "", true
		}
		num := toInt(args[0])
		return strings.ToUpper(fmt.Sprintf("%X", num)), true

	case "oct":
		// OCT(number) - converts number to octal string
		if len(args) == 0 {
			return "", true
		}
		num := toInt(args[0])
		return fmt.Sprintf("%o", num), true

	// Math Functions
	case "abs":
		// ABS(number) - absolute value
		return math.Abs(toFloat(args[0])), true

	case "sqr":
		// SQR(number) - square root
		return math.Sqrt(toFloat(args[0])), true

	case "rnd":
		// RND([seed]) - random number between 0 and 1
		if ctx == nil {
			return rand.Float64(), true
		}
		if len(args) > 0 {
			return ctx.nextRandomValue(args[0], true), true
		}
		return ctx.nextRandomValue(nil, false), true

	case "randomize":
		// RANDOMIZE [seed] - initialize random-number generator
		if ctx == nil {
			return nil, true
		}
		if len(args) > 0 {
			ctx.randomizeWithSeed(seedFromNumber(toFloat(args[0])))
			return nil, true
		}
		ctx.randomizeWithSeed(time.Now().UnixNano())
		return nil, true

	case "round":
		// ROUND(number, [digits]) - rounds to specified digits
		num := toFloat(args[0])
		digits := 0
		if len(args) > 1 {
			digits = toInt(args[1])
		}
		multiplier := math.Pow(10, float64(digits))
		return math.Round(num*multiplier) / multiplier, true

	case "int":
		// INT(number) - integer part (truncates toward negative infinity)
		return int(math.Floor(toFloat(args[0]))), true

	case "fix":
		// FIX(number) - integer part (truncates toward zero)
		f := toFloat(args[0])
		if f >= 0 {
			return int(f), true
		}
		return int(math.Ceil(f)), true

	case "sgn":
		// SGN(number) - sign (-1, 0, 1)
		f := toFloat(args[0])
		if f < 0 {
			return -1, true
		} else if f > 0 {
			return 1, true
		}
		return 0, true

	case "sin":
		// SIN(number) - sine (radians)
		return math.Sin(toFloat(args[0])), true

	case "cos":
		// COS(number) - cosine (radians)
		return math.Cos(toFloat(args[0])), true

	case "tan":
		// TAN(number) - tangent (radians)
		return math.Tan(toFloat(args[0])), true

	case "atn":
		// ATN(number) - arctangent (radians)
		return math.Atan(toFloat(args[0])), true

	case "log":
		// LOG(number) - natural logarithm
		return math.Log(toFloat(args[0])), true

	case "exp":
		// EXP(number) - e raised to power
		return math.Exp(toFloat(args[0])), true

	// Type Conversion Functions
	case "cint":
		// CINT(expression) - convert to integer
		return toInt(args[0]), true

	case "cdbl":
		// CDBL(expression) - convert to double
		return toFloat(args[0]), true

	case "cstr":
		// CSTR(expression) - convert to string
		return toString(args[0]), true

	case "cbool":
		// CBOOL(expression) - convert to boolean
		val := args[0]
		switch v := val.(type) {
		case bool:
			return v, true
		case int:
			return v != 0, true
		case float64:
			return v != 0, true
		case string:
			s := strings.ToLower(strings.TrimSpace(v))
			return s == "true" || s == "-1" || s == "1", true
		default:
			return val != nil, true
		}

	case "cdate":
		// CDATE(expression) - convert to date
		return toDateTime(args[0]), true

	case "cbyte":
		// CBYTE(expression) - convert to byte (0-255)
		n := toInt(args[0])
		n = n % 256
		if n < 0 {
			n = 256 + n
		}
		return n, true

	case "ccur":
		// CCUR(expression) - convert to currency (float64)
		return toFloat(args[0]), true

	case "clng":
		// CLNG(expression) - convert to long integer
		return toInt(args[0]), true

	case "csng":
		// CSNG(expression) - convert to single precision float
		return float32(toFloat(args[0])), true

	// Type Checking Functions (some already handled above)
	case "isnumeric":
		// ISNUMERIC(expression) - checks if string is numeric
		if len(args) == 0 {
			return false, true
		}
		s := toString(args[0])
		_, err := strconv.ParseFloat(s, 64)
		return err == nil, true

	case "isdate":
		// ISDATE(expression) - checks if string is a valid date
		if len(args) == 0 {
			return false, true
		}
		s := toString(args[0])
		// Try parsing with common date formats
		formats := []string{
			"01/02/2006",
			"2006-01-02",
			"January 2, 2006",
			"01/02/2006 15:04:05",
			"2006-01-02 15:04:05",
		}
		for _, format := range formats {
			_, err := time.Parse(format, s)
			if err == nil {
				return true, true
			}
		}
		return false, true

	// Formatting Functions
	case "formatcurrency":
		// FORMATCURRENCY(value, [digits], [include_leading_digit], [use_parens], [group_separator])
		if len(args) == 0 {
			return "", true
		}
		value := toFloat(args[0])
		digits := 2
		if len(args) > 1 {
			digits = toInt(args[1])
		}
		includeLeading := true
		if len(args) > 2 {
			includeLeading = isTruthy(args[2])
		}
		useParens := false
		if len(args) > 3 {
			useParens = isTruthy(args[3])
		}
		groupDigits := false
		if len(args) > 4 {
			groupDigits = isTruthy(args[4])
		}

		return formatNumeric(value, digits, includeLeading, useParens, groupDigits, "$", ""), true

	case "formatnumber":
		// FORMATNUMBER(value, [digits], [include_leading_digit], [use_parens], [group_separator])
		if len(args) == 0 {
			return "", true
		}
		value := toFloat(args[0])
		digits := 0
		if len(args) > 1 {
			digits = toInt(args[1])
		}
		includeLeading := true
		if len(args) > 2 {
			includeLeading = isTruthy(args[2])
		}
		useParens := false
		if len(args) > 3 {
			useParens = isTruthy(args[3])
		}
		groupDigits := false
		if len(args) > 4 {
			groupDigits = isTruthy(args[4])
		}

		return formatNumeric(value, digits, includeLeading, useParens, groupDigits, "", ""), true

	case "formatpercent":
		// FORMATPERCENT(value, [digits], [include_leading_digit], [use_parens], [group_separator])
		if len(args) == 0 {
			return "", true
		}
		value := toFloat(args[0]) * 100
		digits := 0
		if len(args) > 1 {
			digits = toInt(args[1])
		}
		includeLeading := true
		if len(args) > 2 {
			includeLeading = isTruthy(args[2])
		}
		useParens := false
		if len(args) > 3 {
			useParens = isTruthy(args[3])
		}
		groupDigits := false
		if len(args) > 4 {
			groupDigits = isTruthy(args[4])
		}

		return formatNumeric(value, digits, includeLeading, useParens, groupDigits, "", "%"), true

	// Date/Time extra functions
	case "timer":
		// TIMER() - returns seconds since midnight
		now := time.Now()
		midnight := time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, now.Location())
		return now.Sub(midnight).Seconds(), true

	case "weekdayname":
		// WEEKDAYNAME(weekday, [abbreviate]) - returns day name
		if len(args) == 0 {
			return "", true
		}
		w := toInt(args[0])
		if w < 1 || w > 7 {
			return "", true
		}
		// VBScript: 1=Sunday, 7=Saturday
		// Go: 0=Sunday, 6=Saturday
		days := []string{"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}
		return days[w-1], true

	case "monthname":
		// MONTHNAME(month, [abbreviate]) - returns month name
		if len(args) == 0 {
			return "", true
		}
		m := toInt(args[0])
		if m < 1 || m > 12 {
			return "", true
		}
		months := []string{"January", "February", "March", "April", "May", "June",
			"July", "August", "September", "October", "November", "December"}
		return months[m-1], true

	default:
		return nil, false
	}
}

// getTypeName returns the VBScript type name for a value
func getTypeName(val interface{}) string {
	if val == nil {
		return "Empty"
	}

	switch v := val.(type) {
	case EmptyValue:
		return "Empty"
	case NothingValue:
		return "Nothing"
	case bool:
		return "Boolean"
	case int, int8, int16, int32, int64:
		return "Integer"
	case float32, float64:
		return "Double"
	case string:
		return "String"
	case time.Time:
		return "Date"
	case *VBArray:
		return "Variant()"
	case []interface{}:
		return "Variant()"
	case map[string]interface{}:
		return "Dictionary"
	case ASPLibrary:
		return "Object"
	default:
		// Check if it's a struct/custom type
		rv := reflect.ValueOf(v)
		if rv.Kind() == reflect.Ptr {
			rv = rv.Elem()
		}
		if rv.Kind() == reflect.Struct {
			return "Object"
		}
		return "Unknown"
	}
}

// getVarType returns the VBScript VarType constant for a value
// VBScript VarType constants:
// 0 = vbEmpty, 1 = vbNull, 2 = vbInteger, 3 = vbLong, 4 = vbSingle
// 5 = vbDouble, 6 = vbCurrency, 7 = vbDate, 8 = vbString, 9 = vbObject
// 10 = vbError, 11 = vbBoolean, 12 = vbVariant, 13 = vbDataObject
// 17 = vbByte, 8192 = vbArray (flag), 8204 = vbArray + vbVariant
func getVarType(val interface{}) int {
	if val == nil {
		return 0 // vbEmpty
	}

	switch v := val.(type) {
	case EmptyValue:
		return 0 // vbEmpty
	case NothingValue:
		return 1 // vbNull (Nothing is treated as Null for VarType)
	case bool:
		return 11 // vbBoolean
	case int, int8, int16, int32:
		return 2 // vbInteger
	case int64:
		return 3 // vbLong
	case float32:
		return 4 // vbSingle
	case float64:
		return 5 // vbDouble
	case string:
		return 8 // vbString
	case time.Time:
		return 7 // vbDate
	case *VBArray:
		return 8204
	case []interface{}:
		// Array of Variant: 8192 (vbArray flag) + 12 (vbVariant) = 8204
		return 8204
	case map[string]interface{}, ASPLibrary:
		return 9 // vbObject
	default:
		// Check if it's a struct/custom type (treat as object)
		rv := reflect.ValueOf(v)
		if rv.Kind() == reflect.Ptr {
			rv = rv.Elem()
		}
		if rv.Kind() == reflect.Struct {
			return 9 // vbObject
		}
		return 9 // Default to vbObject for unknown types
	}
}

// parseHexLiteral parses hexadecimal literals (&h prefix)
func parseHexLiteral(s string) (int64, bool) {
	s = strings.TrimSpace(s)
	if len(s) < 3 {
		return 0, false
	}

	// Check for &h or &H prefix
	if strings.HasPrefix(strings.ToLower(s), "&h") {
		val, err := strconv.ParseInt(s[2:], 16, 64)
		if err == nil {
			return val, true
		}
	}

	return 0, false
}

// parseOctalLiteral parses octal literals (&o prefix)
func parseOctalLiteral(s string) (int64, bool) {
	s = strings.TrimSpace(s)
	if len(s) < 3 {
		return 0, false
	}

	// Check for &o or &O prefix
	if strings.HasPrefix(strings.ToLower(s), "&o") {
		val, err := strconv.ParseInt(s[2:], 8, 64)
		if err == nil {
			return val, true
		}
	}

	return 0, false
}

// tryParseNumericLiteral attempts to parse numeric literals including hex and octal
func tryParseNumericLiteral(s string) (interface{}, bool) {
	s = strings.TrimSpace(s)

	// Try hexadecimal
	if val, ok := parseHexLiteral(s); ok {
		return int(val), true
	}

	// Try octal
	if val, ok := parseOctalLiteral(s); ok {
		return int(val), true
	}

	// Try decimal integer
	if val, err := strconv.ParseInt(s, 10, 64); err == nil {
		return int(val), true
	}

	// Try floating point
	if val, err := strconv.ParseFloat(s, 64); err == nil {
		return val, true
	}

	return nil, false
}

func formatNumeric(value float64, digits int, includeLeadingDigit, useParens, groupDigits bool, prefix, suffix string) string {
	if digits < 0 {
		digits = 0
	}

	negative := value < 0
	absVal := math.Abs(value)
	formatStr := "%." + strconv.Itoa(digits) + "f"
	formatted := fmt.Sprintf(formatStr, absVal)

	intPart := formatted
	fracPart := ""
	if dot := strings.IndexByte(formatted, '.'); dot >= 0 {
		intPart = formatted[:dot]
		fracPart = formatted[dot+1:]
	}

	if groupDigits {
		intPart = applyGrouping(intPart)
	}

	hasFraction := fracPart != ""
	if !includeLeadingDigit && hasFraction && intPart == "0" {
		intPart = ""
	}

	body := intPart
	if hasFraction {
		if intPart == "" {
			body = "." + fracPart
		} else {
			body = intPart + "." + fracPart
		}
	}

	if body == "" {
		body = "0"
	}

	body = prefix + body + suffix

	if negative {
		if useParens {
			body = "(" + body + ")"
		} else {
			body = "-" + body
		}
	}

	return body
}

func applyGrouping(intPart string) string {
	if len(intPart) <= 3 {
		return intPart
	}

	var b strings.Builder
	first := len(intPart) % 3
	if first == 0 {
		first = 3
	}
	b.WriteString(intPart[:first])
	for i := first; i < len(intPart); i += 3 {
		b.WriteByte(',')
		end := i + 3
		if end > len(intPart) {
			end = len(intPart)
		}
		b.WriteString(intPart[i:end])
	}

	return b.String()
}

// evalExpression evaluates an expression string using the execution context
func evalExpression(exprStr string, ctx *ExecutionContext) interface{} {
	if ctx == nil || exprStr == "" {
		return nil
	}

	// Try to get the executor from the Server object
	if ctx.Server != nil {
		if executor, ok := ctx.Server.GetExecutor().(*ASPExecutor); ok && executor != nil {
			// Create a temporary variable name
			tempVarName := "v_eval_res_" + strconv.FormatInt(time.Now().UnixNano(), 16)

			// Wrap expression in an assignment: temp = expression
			code := tempVarName + " = " + exprStr
			// fmt.Printf("Eval Code: %s\n", code)

			// Parse the code
			// We need to use the parser from the asp package or vbscript-go directly
			// Since we are in server package and can't easily access asp package functions without import cycle if asp imports server (which it likely doesn't, but let's check)
			// Actually server imports asp. So we can use asp.NewASPParser.

			// However, to avoid import cycles if asp imports server (unlikely but possible), let's see.
			// asp/asp_parser.go imports vbscript-go.
			// server/executor.go imports asp.

			// So we can use vbscript-go parser directly here.
			parser := vb.NewParser(code)

			// Defer recovery for parser panics
			defer func() {
				if r := recover(); r != nil {
					fmt.Printf("Eval parse panic: %v\n", r)
				}
			}()

			program := parser.Parse()

			if program == nil {
				// Fallback to simple evaluation if parsing fails or returns nil
				fmt.Println("Eval parse returned nil program")
			} else {
				// Execute the program
				// We need to execute it within the current context
				// ASPExecutor has executeVBProgram but it's private (lowercase)
				// We can use NewASPVisitor and VisitStatement manually

				visitor := NewASPVisitor(ctx, executor)

				// Execute all statements (should be just one assignment)
				for _, stmt := range program.Body {
					if err := visitor.VisitStatement(stmt); err != nil {
						fmt.Printf("Eval execution error: %v\n", err)
						return nil
					}
				}

				// Retrieve the result
				if val, exists := ctx.GetVariable(tempVarName); exists {
					// Clean up
					// ctx.RemoveVariable(tempVarName) // If we had such method
					return val
				} else {
					fmt.Println("Eval variable not found after execution")
				}
			}
		}
	}

	// ... existing fallback logic ...
	exprStr = strings.TrimSpace(exprStr)
	if len(exprStr) >= 2 {
		if exprStr[0] == '"' && exprStr[len(exprStr)-1] == '"' {
			// Handle escaped quotes ""
			inner := exprStr[1 : len(exprStr)-1]
			inner = strings.ReplaceAll(inner, `""`, `"`)
			return inner
		}
	}

	// Try to parse as number
	if val, ok := tryParseNumericLiteral(exprStr); ok {
		return val
	}

	// Check for boolean constants
	exprLower := strings.ToLower(exprStr)
	if exprLower == "true" {
		return true
	}
	if exprLower == "false" {
		return false
	}

	// Check for variable reference
	if ctx != nil {
		varName := strings.ToLower(exprStr)
		if val, ok := ctx.GetVariable(varName); ok {
			return val
		}
	}

	// Return nil if unable to evaluate
	return nil
}

// isBuiltInFunctionName checks if the given name (case-insensitive) is a VBScript built-in function.
// This is used to ensure built-in functions take precedence over class methods with the same name.
func isBuiltInFunctionName(name string) bool {
	nameLower := strings.ToLower(name)
	switch nameLower {
	// Type checking functions
	case "isempty", "isnull", "isnumeric", "isdate", "isarray", "isobject":
		return true
	// Conversion functions
	case "cint", "clng", "cdbl", "csng", "cstr", "cbool", "cdate", "cbyte":
		return true
	// String functions
	case "len", "left", "right", "mid", "trim", "ltrim", "rtrim", "ucase", "lcase":
		return true
	case "instr", "instrrev", "replace", "split", "join", "strcomp", "strreverse":
		return true
	case "space", "string", "asc", "chr", "ascw", "chrw":
		return true
	// Math functions
	case "abs", "sgn", "int", "fix", "round", "sqr", "exp", "log", "sin", "cos", "tan", "atn":
		return true
	case "rnd", "randomize":
		return true
	// Array functions
	case "array", "ubound", "lbound", "filter", "redim":
		return true
	// Date/Time functions
	case "now", "date", "time", "year", "month", "day", "hour", "minute", "second", "weekday":
		return true
	case "dateadd", "datediff", "datepart", "dateserial", "timeserial", "datevalue", "timevalue":
		return true
	case "monthname", "weekdayname", "formatdatetime":
		return true
	// Other common functions
	case "typename", "vartype", "eval", "execute", "executeglobal":
		return true
	case "msgbox", "inputbox":
		return true
	case "hex", "oct":
		return true
	case "getlocale", "setlocale", "formatcurrency", "formatnumber", "formatpercent":
		return true
	case "escape", "unescape":
		return true
	case "getref", "getobject", "createobject":
		return true
	case "rgb":
		return true
	default:
		return false
	}
}
