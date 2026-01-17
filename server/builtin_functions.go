package server

import (
	"reflect"
	"strconv"
	"strings"
)

// EmptyValue represents VBScript Empty type
type EmptyValue struct{}

// NothingValue represents VBScript Nothing type (null object reference)
type NothingValue struct{}

// evalBuiltInFunction evaluates VBScript built-in functions
// Returns (result, wasHandled) where wasHandled indicates if the function was recognized
func evalBuiltInFunction(funcName string, args []interface{}, ctx *ExecutionContext) (interface{}, bool) {
	funcLower := strings.ToLower(funcName)

	// Try date/time functions first
	if result, handled := evalDateTimeFunction(funcLower, args, ctx); handled {
		return result, true
	}

	switch funcLower {
	// Type checking functions
	case "isempty":
		if len(args) == 0 {
			return true, true
		}
		val := args[0]
		// Check if value is Empty (nil or EmptyValue)
		if val == nil {
			return true, true
		}
		if _, ok := val.(EmptyValue); ok {
			return true, true
		}
		// Check if it's a newly declared variable (stored as nil in context)
		return false, true

	case "isnull":
		if len(args) == 0 {
			return false, true
		}
		val := args[0]
		// In VBScript, Null is a special value indicating no valid data
		// We represent it as nil
		return val == nil, true

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
		case map[string]interface{}, []interface{}, ASPLibrary:
			return true, true
		default:
			// Check if it's a struct (custom type)
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
		return result, true

	case "isarray":
		// IsArray(var) - checks if variable is an array
		if len(args) == 0 {
			return false, true
		}
		val := args[0]
		_, ok := val.([]interface{})
		return ok, true

	case "lbound":
		// LBound(array, [dimension]) - returns lower bound (always 0 in VBScript)
		return 0, true

	case "ubound":
		// UBound(array, [dimension]) - returns upper bound (last index)
		if len(args) == 0 {
			return -1, true
		}
		arr, ok := args[0].([]interface{})
		if !ok {
			return -1, true
		}
		// Check for dimension parameter (default is 1)
		dim := 1
		if len(args) > 1 {
			dim = toInt(args[1])
		}
		// For dimension 1, return length-1
		if dim == 1 {
			return len(arr) - 1, true
		}
		// For nested arrays, traverse to requested dimension
		current := interface{}(arr)
		for d := 1; d <= dim; d++ {
			if slice, ok := current.([]interface{}); ok {
				if d == dim {
					return len(slice) - 1, true
				}
				if len(slice) > 0 {
					current = slice[0]
					continue
				}
				return -1, true
			}
			return -1, true
		}
		return -1, true

	case "split":
		// Split(expression, delimiter, [limit], [compare]) - splits string into array
		if len(args) < 2 {
			return []interface{}{}, true
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
		return result, true

	case "join":
		// Join(array, delimiter) - joins array elements into string
		if len(args) < 2 {
			return "", true
		}
		arr, ok := args[0].([]interface{})
		if !ok {
			return "", true
		}
		delim := toString(args[1])
		
		parts := make([]string, len(arr))
		for i, item := range arr {
			parts[i] = toString(item)
		}
		return strings.Join(parts, delim), true
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
