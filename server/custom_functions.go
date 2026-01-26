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
	"crypto/md5"
	"crypto/sha1"
	"crypto/sha256"
	"encoding/base64"
	"encoding/csv"
	"fmt"
	"html"
	"math"
	"math/rand"
	"net"
	"net/mail"
	"net/url"
	"regexp"
	"strings"
	"time"
)

// arrayValues extracts the raw slice from VBArray or []interface{} values.
func arrayValues(val interface{}) ([]interface{}, bool) {
	if arr, ok := toVBArray(val); ok {
		return arr.Values, true
	}
	return nil, false
}

// evalCustomFunction evaluates custom G3 functions following VBScript conventions
// Returns (result, wasHandled) where wasHandled indicates if the function was recognized
func evalCustomFunction(funcName string, args []interface{}, ctx *ExecutionContext) (interface{}, bool) {
	funcLower := strings.ToLower(funcName)

	switch funcLower {
	// Document functions
	case "document.write", "documentwrite":
		if len(args) == 0 {
			return nil, true
		}
		val := toString(args[0])
		// HTML encode for safety
		encoded := html.EscapeString(val)
		ctx.Response.Write(encoded)
		return nil, true

	// Array functions
	case "axarraymerge":
		// Accepts multiple arrays and returns merged array
		result := []interface{}{}
		for _, arg := range args {
			if arr, ok := arrayValues(arg); ok {
				result = append(result, arr...)
			} else if dict, ok := arg.(map[string]interface{}); ok {
				for _, v := range dict {
					result = append(result, v)
				}
			}
		}
		return NewVBArrayFromValues(0, result), true

	case "axarraycontains":
		// in_array equivalent: search for value in array
		if len(args) < 2 {
			return false, true
		}
		needle := args[0]
		switch haystack := args[1].(type) {
		case []interface{}:
			for _, v := range haystack {
				if areEqual(needle, v) {
					return true, true
				}
			}
		case map[string]interface{}:
			for _, v := range haystack {
				if areEqual(needle, v) {
					return true, true
				}
			}
		default:
			if arr, ok := arrayValues(args[1]); ok {
				for _, v := range arr {
					if areEqual(needle, v) {
						return true, true
					}
				}
			}
		}
		return false, true

	case "axarraymap":
		// array_map equivalent: transform array elements using callback
		if len(args) < 2 {
			return NewVBArrayFromValues(0, []interface{}{}), true
		}
		callbackName := toString(args[0])
		result := []interface{}{}

		switch arr := args[1].(type) {
		case []interface{}:
			for _, item := range arr {
				if res, handled := evalBuiltInFunction(callbackName, []interface{}{item}, ctx); handled {
					result = append(result, res)
				}
			}
		case map[string]interface{}:
			for _, item := range arr {
				if res, handled := evalBuiltInFunction(callbackName, []interface{}{item}, ctx); handled {
					result = append(result, res)
				}
			}
		default:
			if arrVals, ok := arrayValues(args[1]); ok {
				for _, item := range arrVals {
					if res, handled := evalBuiltInFunction(callbackName, []interface{}{item}, ctx); handled {
						result = append(result, res)
					}
				}
			}
		}
		return NewVBArrayFromValues(0, result), true

	case "axarrayfilter":
		// array_filter equivalent: filter array using callback
		if len(args) < 2 {
			return NewVBArrayFromValues(0, []interface{}{}), true
		}
		callbackName := toString(args[0])
		result := []interface{}{}

		switch arr := args[1].(type) {
		case []interface{}:
			for _, item := range arr {
				if res, handled := evalBuiltInFunction(callbackName, []interface{}{item}, ctx); handled {
					if isTruthyCustom(res) {
						result = append(result, item)
					}
				}
			}
		case map[string]interface{}:
			for _, item := range arr {
				if res, handled := evalBuiltInFunction(callbackName, []interface{}{item}, ctx); handled {
					if isTruthyCustom(res) {
						result = append(result, item)
					}
				}
			}
		default:
			if arrVals, ok := arrayValues(args[1]); ok {
				for _, item := range arrVals {
					if res, handled := evalBuiltInFunction(callbackName, []interface{}{item}, ctx); handled {
						if isTruthyCustom(res) {
							result = append(result, item)
						}
					}
				}
			}
		}
		return NewVBArrayFromValues(0, result), true

	case "axcount":
		// count equivalent: return array/collection length
		if len(args) == 0 {
			return 0, true
		}
		switch v := args[0].(type) {
		case []interface{}:
			return len(v), true
		case map[string]interface{}:
			return len(v), true
		case string:
			return len(v), true
		case nil, EmptyValue:
			return 0, true
		}
		return 0, true

	case "axarrayreverse":
		// array_reverse: reverse array order
		if len(args) == 0 {
			return []interface{}{}, true
		}
		switch arr := args[0].(type) {
		case []interface{}:
			result := make([]interface{}, len(arr))
			for i, v := range arr {
				result[len(arr)-1-i] = v
			}
			return result, true
		}
		return args[0], true

	case "axrange":
		// range: create array of values from start to end
		if len(args) < 2 {
			return []interface{}{}, true
		}
		start := toInt(args[0])
		end := toInt(args[1])
		step := 1
		if len(args) > 2 {
			step = toInt(args[2])
		}
		if step == 0 {
			step = 1
		}

		result := []interface{}{}
		if step > 0 {
			for i := start; i <= end; i += step {
				result = append(result, i)
			}
		} else {
			for i := start; i >= end; i += step {
				result = append(result, i)
			}
		}
		return result, true

	case "aximplode":
		// implode/join: join array elements with glue
		if len(args) < 2 {
			return "", true
		}
		glue := toString(args[0])
		var pieces []string

		switch arr := args[1].(type) {
		case []interface{}:
			for _, v := range arr {
				pieces = append(pieces, toString(v))
			}
		case map[string]interface{}:
			for _, v := range arr {
				pieces = append(pieces, toString(v))
			}
		}
		return strings.Join(pieces, glue), true

	// String functions
	case "axexplode":
		// explode: split string by delimiter
		if len(args) < 2 {
			return NewVBArrayFromValues(0, []interface{}{}), true
		}
		delimiter := toString(args[0])
		str := toString(args[1])
		limit := -1
		if len(args) > 2 {
			limit = toInt(args[2])
		}

		var parts []string
		if delimiter == "" {
			parts = strings.Split(str, "")
		} else {
			parts = strings.Split(str, delimiter)
		}

		if limit > 0 && len(parts) > limit {
			parts = parts[:limit]
		}

		result := make([]interface{}, len(parts))
		for i, p := range parts {
			result[i] = p
		}
		return NewVBArrayFromValues(0, result), true

	case "axstringreplace":
		// str_replace: replace search string with replacement
		if len(args) < 3 {
			return "", true
		}
		search := args[0]
		replace := args[1]
		subject := toString(args[2])

		// Handle array search/replace
		if searchArr, ok := arrayValues(search); ok {
			replaceArr, _ := arrayValues(replace)
			for i, searchStr := range searchArr {
				replaceStr := ""
				if i < len(replaceArr) {
					replaceStr = toString(replaceArr[i])
				} else {
					replaceStr = toString(replace)
				}
				subject = strings.ReplaceAll(subject, toString(searchStr), replaceStr)
			}
		} else {
			searchStr := toString(search)
			replaceStr := toString(replace)
			subject = strings.ReplaceAll(subject, searchStr, replaceStr)
		}
		return subject, true

	case "axsprintf":
		// sprintf: format string like C printf
		if len(args) == 0 {
			return "", true
		}
		format := toString(args[0])
		argValues := args[1:]
		return formatString(format, argValues), true

	case "axpad":
		// str_pad: pad string to length
		if len(args) < 2 {
			return "", true
		}
		str := toString(args[0])
		length := toInt(args[1])
		padString := " "
		padType := 1 // STR_PAD_RIGHT = 1

		if len(args) > 2 {
			padString = toString(args[2])
		}
		if len(args) > 3 {
			padType = toInt(args[3])
		}

		if len(str) >= length {
			return str, true
		}

		padLen := length - len(str)
		padding := ""
		for len(padding) < padLen {
			padding += padString
		}
		padding = padding[:padLen]

		switch padType {
		case 0: // STR_PAD_LEFT
			return padding + str, true
		case 2: // STR_PAD_BOTH
			leftPad := padLen / 2
			rightPad := padLen - leftPad
			leftPadding := ""
			rightPadding := ""
			for len(leftPadding) < leftPad {
				leftPadding += padString
			}
			leftPadding = leftPadding[:leftPad]
			for len(rightPadding) < rightPad {
				rightPadding += padString
			}
			rightPadding = rightPadding[:rightPad]
			return leftPadding + str + rightPadding, true
		default: // STR_PAD_RIGHT = 1
			return str + padding, true
		}

	case "axrepeat":
		// str_repeat: repeat string
		if len(args) < 2 {
			return "", true
		}
		str := toString(args[0])
		times := toInt(args[1])
		if times < 0 {
			times = 0
		}
		return strings.Repeat(str, times), true

	case "axucfirst":
		// ucfirst: uppercase first character
		if len(args) == 0 {
			return "", true
		}
		str := toString(args[0])
		if len(str) == 0 {
			return str, true
		}
		return strings.ToUpper(str[:1]) + str[1:], true

	case "axwordcount":
		// str_word_count: count words in string
		if len(args) == 0 {
			return 0, true
		}
		str := toString(args[0])
		format := 0 // default: return number of words
		if len(args) > 1 {
			format = toInt(args[1])
		}

		// Split by whitespace
		words := strings.Fields(str)
		count := len(words)

		switch format {
		case 0:
			return count, true
		case 1:
			// Return array of words
			result := make([]interface{}, len(words))
			for i, w := range words {
				result[i] = w
			}
			return result, true
		}
		return count, true

	case "axnl2br":
		// nl2br: convert newlines to <br>
		if len(args) == 0 {
			return "", true
		}
		str := toString(args[0])
		// Replace common newline patterns
		str = strings.ReplaceAll(str, "\r\n", "<br>")
		str = strings.ReplaceAll(str, "\n", "<br>")
		str = strings.ReplaceAll(str, "\r", "<br>")
		return str, true

	case "axtrim":
		// trim with custom characters
		if len(args) == 0 {
			return "", true
		}
		str := toString(args[0])
		chars := " \t\n\r\v\f"
		if len(args) > 1 {
			chars = toString(args[1])
		}
		return strings.Trim(str, chars), true

	case "axstringgetcsv":
		// str_getcsv: parse CSV string
		if len(args) == 0 {
			return []interface{}{}, true
		}
		str := toString(args[0])
		delimiter := ","

		if len(args) > 1 {
			delimiter = toString(args[1])
		}

		reader := csv.NewReader(strings.NewReader(str))
		if len(delimiter) > 0 {
			reader.Comma = rune(delimiter[0])
		}

		record, err := reader.Read()
		if err != nil {
			return []interface{}{}, true
		}

		result := make([]interface{}, len(record))
		for i, v := range record {
			result[i] = v
		}
		return result, true

	// Math functions
	case "axceil":
		// ceil: round up
		if len(args) == 0 {
			return 0, true
		}
		return math.Ceil(toFloat(args[0])), true

	case "axfloor":
		// floor: round down
		if len(args) == 0 {
			return 0, true
		}
		return math.Floor(toFloat(args[0])), true

	case "axmax":
		// max: return maximum value
		if len(args) == 0 {
			return 0, true
		}
		max := toFloat(args[0])
		for _, arg := range args[1:] {
			v := toFloat(arg)
			if v > max {
				max = v
			}
		}
		return max, true

	case "axmin":
		// min: return minimum value
		if len(args) == 0 {
			return 0, true
		}
		min := toFloat(args[0])
		for _, arg := range args[1:] {
			v := toFloat(arg)
			if v < min {
				min = v
			}
		}
		return min, true

	case "axrand":
		// rand: random integer
		if len(args) == 0 {
			return rand.Int(), true
		}
		if len(args) == 1 {
			max := toInt(args[0])
			return rand.Intn(max + 1), true
		}
		min := toInt(args[0])
		max := toInt(args[1])
		if min > max {
			min, max = max, min
		}
		return min + rand.Intn(max-min+1), true

	case "axnumberformat":
		// number_format: format number
		if len(args) == 0 {
			return "", true
		}
		num := toFloat(args[0])
		decimals := 0
		decPoint := "."
		thousandsSep := ","

		if len(args) > 1 {
			decimals = toInt(args[1])
		}
		if len(args) > 2 {
			decPoint = toString(args[2])
		}
		if len(args) > 3 {
			thousandsSep = toString(args[3])
		}

		formatted := fmt.Sprintf("%.*f", decimals, num)
		parts := strings.Split(formatted, ".")

		// Add thousands separator
		intPart := parts[0]
		if thousandsSep != "" {
			var result string
			for i, ch := range intPart {
				if i > 0 && (len(intPart)-i)%3 == 0 && ch != '-' {
					result += thousandsSep
				}
				result += string(ch)
			}
			intPart = result
		}

		if decimals > 0 && len(parts) > 1 {
			return intPart + decPoint + parts[1], true
		}
		return intPart, true

	case "axpi":
		// pi: return the mathematical constant pi
		return math.Pi, true

	// Type checking functions
	case "axisint":
		// is_int: check if value is integer
		if len(args) == 0 {
			return false, true
		}
		_, ok := args[0].(int)
		if !ok {
			_, ok = args[0].(int64)
		}
		return ok, true

	case "axisfloat":
		// is_float: check if value is float
		if len(args) == 0 {
			return false, true
		}
		_, ok := args[0].(float64)
		return ok, true

	case "axctypealpha":
		// ctype_alpha: check if all characters are alphabetic
		if len(args) == 0 {
			return false, true
		}
		str := toString(args[0])
		if str == "" {
			return false, true
		}
		for _, ch := range str {
			if !((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z')) {
				return false, true
			}
		}
		return true, true

	case "axctypealnum":
		// ctype_alnum: check if all characters are alphanumeric
		if len(args) == 0 {
			return false, true
		}
		str := toString(args[0])
		if str == "" {
			return false, true
		}
		for _, ch := range str {
			if !((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9')) {
				return false, true
			}
		}
		return true, true

	case "axempty":
		// empty: check if value is empty
		if len(args) == 0 {
			return true, true
		}
		return isEmpty(args[0]), true

	case "axisset":
		// isset: check if variable is set (not null/empty)
		if len(args) == 0 {
			return false, true
		}
		val := args[0]
		if val == nil {
			return false, true
		}
		if _, ok := val.(EmptyValue); ok {
			return false, true
		}
		return true, true

	// Date/Time functions
	case "axtime":
		// time: return current Unix timestamp
		return time.Now().Unix(), true

	case "axdate":
		// date: format date/time
		if len(args) == 0 {
			return "", true
		}
		format := toString(args[0])
		timestamp := time.Now().Unix()
		if len(args) > 1 {
			timestamp = int64(toInt(args[1]))
		}

		t := time.Unix(timestamp, 0)
		return formatDate(format, t), true

	// Encoding/Hashing functions
	case "axmd5":
		// md5: return MD5 hash
		if len(args) == 0 {
			return "", true
		}
		str := toString(args[0])
		return fmt.Sprintf("%x", md5.Sum([]byte(str))), true

	case "axsha1":
		// sha1: return SHA1 hash
		if len(args) == 0 {
			return "", true
		}
		str := toString(args[0])
		return fmt.Sprintf("%x", sha1.Sum([]byte(str))), true

	case "axhash":
		// hash: return hash of string
		if len(args) < 2 {
			return "", true
		}
		algo := strings.ToLower(toString(args[0]))
		str := toString(args[1])

		switch algo {
		case "sha256":
			return fmt.Sprintf("%x", sha256.Sum256([]byte(str))), true
		case "sha1":
			return fmt.Sprintf("%x", sha1.Sum([]byte(str))), true
		case "md5":
			return fmt.Sprintf("%x", md5.Sum([]byte(str))), true
		}
		return "", true

	case "axbase64encode":
		// base64_encode: encode string to base64
		if len(args) == 0 {
			return "", true
		}
		str := toString(args[0])
		return base64.StdEncoding.EncodeToString([]byte(str)), true

	case "axbase64decode":
		// base64_decode: decode base64 string
		if len(args) == 0 {
			return "", true
		}
		str := toString(args[0])
		decoded, err := base64.StdEncoding.DecodeString(str)
		if err != nil {
			return "", true
		}
		return string(decoded), true

	case "axurldecode":
		// urldecode: decode URL-encoded string
		if len(args) == 0 {
			return "", true
		}
		str := toString(args[0])
		decoded, err := url.QueryUnescape(str)
		if err != nil {
			return str, true
		}
		return decoded, true

	case "axrawurldecode":
		// rawurldecode: decode raw URL-encoded string
		if len(args) == 0 {
			return "", true
		}
		str := toString(args[0])
		// Replace + with space first
		str = strings.ReplaceAll(str, "+", " ")
		decoded, err := url.QueryUnescape(str)
		if err != nil {
			return str, true
		}
		return decoded, true

	case "axrgbtohex":
		// RGBToHex: convert RGB to hex color
		if len(args) < 3 {
			return "#000000", true
		}
		r := toInt(args[0]) & 0xFF
		g := toInt(args[1]) & 0xFF
		b := toInt(args[2]) & 0xFF
		return fmt.Sprintf("#%02X%02X%02X", r, g, b), true

	case "axhtmlspecialchars":
		// htmlspecialchars: encode HTML special characters
		if len(args) == 0 {
			return "", true
		}
		str := toString(args[0])
		return html.EscapeString(str), true

	case "axstriptags":
		// strip_tags: remove HTML/PHP tags
		if len(args) == 0 {
			return "", true
		}
		str := toString(args[0])

		// Simple regex-based tag removal
		re := regexp.MustCompile(`<[^>]*>`)
		result := re.ReplaceAllString(str, "")
		return result, true

	// Validation functions
	case "axfiltervalidateip":
		// filter_var with FILTER_VALIDATE_IP
		if len(args) == 0 {
			return false, true
		}
		ip := toString(args[0])
		ipObj := net.ParseIP(ip)
		return ipObj != nil, true

	case "axfiltervalidateemail":
		// filter_var with FILTER_VALIDATE_EMAIL
		if len(args) == 0 {
			return false, true
		}
		email := toString(args[0])
		_, err := mail.ParseAddress(email)
		return err == nil, true

	// Request arrays
	case "axgetrequest":
		// Get all request parameters (merged GET and POST)
		result := NewDictionary(ctx)
		if ctx == nil || ctx.Request == nil {
			return result, true
		}
		// Merge QueryString parameters
		if ctx.Request.QueryString != nil {
			for _, key := range ctx.Request.QueryString.GetKeys() {
				result.SetProperty(key, ctx.Request.QueryString.Get(key))
			}
		}
		// Merge Form parameters (POST overwrites GET)
		if ctx.Request.Form != nil {
			for _, key := range ctx.Request.Form.GetKeys() {
				result.SetProperty(key, ctx.Request.Form.Get(key))
			}
		}
		return result, true

	case "axgetget":
		// Get all GET parameters
		result := NewDictionary(ctx)
		if ctx == nil || ctx.Request == nil {
			return result, true
		}
		if ctx.Request.QueryString != nil {
			for _, key := range ctx.Request.QueryString.GetKeys() {
				result.SetProperty(key, ctx.Request.QueryString.Get(key))
			}
		}
		return result, true

	case "axgetpost":
		// Get all POST parameters
		result := NewDictionary(ctx)
		if ctx == nil || ctx.Request == nil {
			return result, true
		}
		if ctx.Request.Form != nil {
			for _, key := range ctx.Request.Form.GetKeys() {
				result.SetProperty(key, ctx.Request.Form.Get(key))
			}
		}
		return result, true

	// Utility functions
	case "axvardump":
		// var_dump: dump variable information
		if len(args) == 0 {
			return nil, true
		}
		output := dumpVariable(args[0], 0)
		ctx.Response.Write(output)
		return nil, true

	case "axgenerateguid", "generateguid":
		// Generate GUID - unique identifier
		return generateUUID(), true

	case "axbuildquerystring", "buildquerystring":
		// Build URL query string from dictionary
		if len(args) == 0 {
			return "", true
		}
		switch dict := args[0].(type) {
		case map[string]interface{}:
			values := url.Values{}
			for k, v := range dict {
				key := strings.ToLower(k)
				values.Set(key, toString(v))
			}
			return values.Encode(), true
		case *DictionaryLibrary:
			// Handle Scripting.Dictionary object (wrapped as DictionaryLibrary)
			values := url.Values{}
			keys, _ := dict.CallMethod("keys")
			if keyArr, ok := arrayValues(keys); ok {
				for _, k := range keyArr {
					keyStr := toString(k)
					val, _ := dict.CallMethod("item", k)
					values.Set(strings.ToLower(keyStr), toString(val))
				}
			}
			return values.Encode(), true
		case *Dictionary:
			// Handle raw Dictionary object
			values := url.Values{}
			keys := dict.Keys([]interface{}{})
			if keyArr, ok := arrayValues(keys); ok {
				for _, k := range keyArr {
					keyStr := toString(k)
					val := dict.Item([]interface{}{k})
					values.Set(strings.ToLower(keyStr), toString(val))
				}
			}
			return values.Encode(), true
		case []interface{}:
			// Handle array of key-value pairs
			values := url.Values{}
			for i := 0; i < len(dict)-1; i += 2 {
				key := toString(dict[i])
				val := toString(dict[i+1])
				values.Set(strings.ToLower(key), val)
			}
			return values.Encode(), true
		}

		if arr, ok := arrayValues(args[0]); ok {
			values := url.Values{}
			for i := 0; i+1 < len(arr); i += 2 {
				key := toString(arr[i])
				val := toString(arr[i+1])
				values.Set(strings.ToLower(key), val)
			}
			return values.Encode(), true
		}
		return "", true
	}

	return nil, false
}

// Helper functions

func areEqual(a, b interface{}) bool {
	if a == b {
		return true
	}
	aStr := toString(a)
	bStr := toString(b)
	return aStr == bStr
}

func isTruthyCustom(v interface{}) bool {
	if v == nil {
		return false
	}
	if _, ok := v.(EmptyValue); ok {
		return false
	}
	if b, ok := v.(bool); ok {
		return b
	}
	if s, ok := v.(string); ok {
		return s != ""
	}
	if i, ok := v.(int); ok {
		return i != 0
	}
	if i, ok := v.(int64); ok {
		return i != 0
	}
	if f, ok := v.(float64); ok {
		return f != 0
	}
	return true
}

func isEmpty(v interface{}) bool {
	if v == nil {
		return true
	}
	if _, ok := v.(EmptyValue); ok {
		return true
	}
	if s, ok := v.(string); ok {
		return s == ""
	}
	if i, ok := v.(int); ok {
		return i == 0
	}
	if i, ok := v.(int64); ok {
		return i == 0
	}
	if f, ok := v.(float64); ok {
		return f == 0
	}
	if b, ok := v.(bool); ok {
		return !b
	}
	if arr, ok := arrayValues(v); ok {
		return len(arr) == 0
	}
	if dict, ok := v.(map[string]interface{}); ok {
		return len(dict) == 0
	}
	return false
}

func formatString(format string, args []interface{}) string {
	result := format
	argIndex := 0

	// Simple sprintf implementation
	for i := 0; i < len(result); i++ {
		if result[i] == '%' && i+1 < len(result) {
			if argIndex >= len(args) {
				break
			}

			spec := result[i+1]
			var replacement string

			switch spec {
			case 's':
				replacement = toString(args[argIndex])
			case 'd', 'u':
				replacement = fmt.Sprintf("%d", toInt(args[argIndex]))
			case 'f':
				replacement = fmt.Sprintf("%f", toFloat(args[argIndex]))
			case 'x':
				replacement = fmt.Sprintf("%x", toInt(args[argIndex]))
			case 'X':
				replacement = fmt.Sprintf("%X", toInt(args[argIndex]))
			case '%':
				replacement = "%"
				argIndex--
			default:
				replacement = ""
				argIndex--
			}

			result = result[:i] + replacement + result[i+2:]
			i += len(replacement) - 1
			argIndex++
		}
	}

	return result
}

func formatDate(format string, t time.Time) string {
	result := format
	replacements := map[string]string{
		"Y": fmt.Sprintf("%d", t.Year()),
		"y": fmt.Sprintf("%02d", t.Year()%100),
		"m": fmt.Sprintf("%02d", t.Month()),
		"n": fmt.Sprintf("%d", t.Month()),
		"d": fmt.Sprintf("%02d", t.Day()),
		"j": fmt.Sprintf("%d", t.Day()),
		"H": fmt.Sprintf("%02d", t.Hour()),
		"i": fmt.Sprintf("%02d", t.Minute()),
		"s": fmt.Sprintf("%02d", t.Second()),
		"w": fmt.Sprintf("%d", t.Weekday()),
		"z": fmt.Sprintf("%d", t.YearDay()-1),
		"W": fmt.Sprintf("%02d", getWeekNumber(t)),
		"F": t.Month().String(),
		"M": t.Month().String()[:3],
		"l": t.Weekday().String(),
		"D": t.Weekday().String()[:3],
	}

	for key, val := range replacements {
		result = strings.ReplaceAll(result, key, val)
	}

	return result
}

func getWeekNumber(t time.Time) int {
	_, week := t.ISOWeek()
	return week
}

func dumpVariable(v interface{}, depth int) string {
	indent := strings.Repeat("  ", depth)
	switch val := v.(type) {
	case nil:
		return "NULL"
	case EmptyValue:
		return "Empty"
	case bool:
		if val {
			return "bool(true)"
		}
		return "bool(false)"
	case int:
		return fmt.Sprintf("int(%d)", val)
	case int64:
		return fmt.Sprintf("int(%d)", val)
	case float64:
		return fmt.Sprintf("float(%f)", val)
	case string:
		return fmt.Sprintf("string(%d) \"%s\"", len(val), val)
	case []interface{}:
		result := fmt.Sprintf("array(%d) {\n", len(val))
		for i, item := range val {
			result += indent + fmt.Sprintf("  [%d]=>\n", i)
			result += indent + "  " + dumpVariable(item, depth+2) + "\n"
		}
		result += indent + "}"
		return result
	case map[string]interface{}:
		result := fmt.Sprintf("array(%d) {\n", len(val))
		for key, item := range val {
			result += indent + fmt.Sprintf("  [\"%s\"]=>\n", key)
			result += indent + "  " + dumpVariable(item, depth+2) + "\n"
		}
		result += indent + "}"
		return result
	default:
		return fmt.Sprintf("object(%T)", v)
	}
}

func generateUUID() string {
	b := make([]byte, 16)
	rand.Read(b)
	b[6] = (b[6] & 0x0f) | 0x40
	b[8] = (b[8] & 0x3f) | 0x80
	return fmt.Sprintf("%x-%x-%x-%x-%x", b[0:4], b[4:6], b[6:8], b[8:10], b[10:])
}

// DocumentObject represents the Document object in ASP (similar to Response but with HTML encoding)
type DocumentObject struct {
	ctx *ExecutionContext
}

// NewDocumentObject creates a new Document object
func NewDocumentObject(ctx *ExecutionContext) *DocumentObject {
	return &DocumentObject{ctx: ctx}
}

// GetName returns the object name
func (d *DocumentObject) GetName() string {
	return "Document"
}

// GetProperty retrieves a property
func (d *DocumentObject) GetProperty(name string) interface{} {
	return nil
}

// SetProperty sets a property
func (d *DocumentObject) SetProperty(name string, value interface{}) error {
	return nil
}

// CallMethod calls a method on Document object
func (d *DocumentObject) CallMethod(name string, args ...interface{}) (interface{}, error) {
	nameLower := strings.ToLower(name)

	switch nameLower {
	case "write":
		if len(args) == 0 {
			return nil, nil
		}
		val := toString(args[0])
		// HTML encode for safety
		encoded := html.EscapeString(val)
		d.ctx.Response.Write(encoded)
		return nil, nil
	}

	return nil, fmt.Errorf("Document.%s is not supported", name)
}
