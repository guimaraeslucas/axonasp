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
	"math"
	"strconv"
	"strings"
	"time"
)

// jsIntlDateTimeFormatObject stores normalized Intl.DateTimeFormat state.
type jsIntlDateTimeFormatObject struct {
	localeTag string
	layout    string
}

// jsIntlNumberFormatObject stores normalized Intl.NumberFormat state.
type jsIntlNumberFormatObject struct {
	localeTag       string
	style           string
	digits          int
	useGrouping     bool
	currencySymbol  string
	currencySpacing string
}

// jsIntlCollatorObject stores normalized Intl.Collator state.
type jsIntlCollatorObject struct {
	localeTag   string
	usage       string
	sensitivity string
	ignorePunct bool
	numeric     bool
	caseFirst   string
}

// jsIntlPluralRulesObject stores normalized Intl.PluralRules state.
type jsIntlPluralRulesObject struct {
	localeTag string
	style     string // "cardinal" (default) or "ordinal"
}

// jsIntlRelativeTimeFormatObject stores normalized Intl.RelativeTimeFormat state.
type jsIntlRelativeTimeFormatObject struct {
	localeTag string
	numeric   string // "always" (default) or "auto"
	style     string // "long" (default), "short", or "narrow"
}

// jsCreateIntlObject allocates the global Intl namespace and its constructor entries.
func (vm *VM) jsCreateIntlObject() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 6)
	obj["__js_type"] = NewString("Intl")
	obj["DateTimeFormat"] = vm.jsCreateIntlDateTimeFormatConstructor()
	obj["NumberFormat"] = vm.jsCreateIntlNumberFormatConstructor()
	obj["Collator"] = vm.jsCreateIntlCollatorConstructor()
	obj["PluralRules"] = vm.jsCreateIntlPluralRulesConstructor()
	obj["RelativeTimeFormat"] = vm.jsCreateIntlRelativeTimeFormatConstructor()
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 7)
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateIntlDateTimeFormatConstructor allocates the Intl.DateTimeFormat constructor object.
func (vm *VM) jsCreateIntlDateTimeFormatConstructor() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 3)
	obj["__js_type"] = NewString("Intl.DateTimeFormat")
	obj["__js_ctor"] = NewString("IntlDateTimeFormat")
	obj["name"] = NewString("DateTimeFormat")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateIntlNumberFormatConstructor allocates the Intl.NumberFormat constructor object.
func (vm *VM) jsCreateIntlNumberFormatConstructor() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 3)
	obj["__js_type"] = NewString("Intl.NumberFormat")
	obj["__js_ctor"] = NewString("IntlNumberFormat")
	obj["name"] = NewString("NumberFormat")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateIntlCollatorConstructor allocates the Intl.Collator constructor object.
func (vm *VM) jsCreateIntlCollatorConstructor() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 3)
	obj["__js_type"] = NewString("Intl.Collator")
	obj["__js_ctor"] = NewString("IntlCollator")
	obj["name"] = NewString("Collator")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateIntlPluralRulesConstructor allocates the Intl.PluralRules constructor object.
func (vm *VM) jsCreateIntlPluralRulesConstructor() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 3)
	obj["__js_type"] = NewString("Intl.PluralRules")
	obj["__js_ctor"] = NewString("IntlPluralRules")
	obj["name"] = NewString("PluralRules")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	return Value{Type: VTJSObject, Num: objID}
}

// jsCreateIntlRelativeTimeFormatConstructor allocates the Intl.RelativeTimeFormat constructor object.
func (vm *VM) jsCreateIntlRelativeTimeFormatConstructor() Value {
	objID := vm.allocJSID()
	obj := make(map[string]Value, 3)
	obj["__js_type"] = NewString("Intl.RelativeTimeFormat")
	obj["__js_ctor"] = NewString("IntlRelativeTimeFormat")
	obj["name"] = NewString("RelativeTimeFormat")
	vm.jsObjectItems[objID] = obj
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 4)
	return Value{Type: VTJSObject, Num: objID}
}

// jsIntlCreateMethodFunction allocates one bound method for an Intl instance.
func (vm *VM) jsIntlCreateMethodFunction(ownerID int64, methodName string, methodCtor string) Value {
	fn := vm.jsCreateIntrinsicFunction(methodName, methodCtor)
	if fn.Type == VTJSFunction {
		if obj, ok := vm.jsObjectItems[fn.Num]; ok {
			obj["__js_intl_owner"] = NewInteger(ownerID)
		}
	}
	return fn
}

// jsIntlCreateDateTimeFormat allocates one Intl.DateTimeFormat instance with normalized locale state.
func (vm *VM) jsIntlCreateDateTimeFormat(args []Value) Value {
	profile, localeTag := vm.jsIntlResolveLocaleProfile(jsArgOrUndefined(args, 0))
	options := Value{Type: VTJSUndefined}
	if len(args) > 1 {
		options = args[1]
	}
	layout := vm.jsIntlResolveDateTimeLayout(profile, options)
	objID := vm.allocJSID()
	obj := make(map[string]Value, 6)
	obj["__js_type"] = NewString("Intl.DateTimeFormat")
	obj["__js_ctor"] = NewString("IntlDateTimeFormat")
	obj["__js_intl_locale"] = NewString(localeTag)
	obj["__js_intl_layout"] = NewString(layout)
	obj["format"] = vm.jsIntlCreateMethodFunction(objID, "format", "IntlDateTimeFormatFormat")
	obj["formatToParts"] = vm.jsIntlCreateMethodFunction(objID, "formatToParts", "IntlDateTimeFormatFormatToParts")
	vm.jsObjectItems[objID] = obj
	vm.jsIntlDateTimeFormatItems[objID] = &jsIntlDateTimeFormatObject{localeTag: localeTag, layout: layout}
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 7)
	vm.jsSetDescriptor(objID, "format", jsPropertyDescriptor{
		Value:        obj["format"],
		HasValue:     true,
		Enumerable:   false,
		Configurable: true,
		Writable:     true,
	})
	vm.jsSetDescriptor(objID, "formatToParts", jsPropertyDescriptor{
		Value:        obj["formatToParts"],
		HasValue:     true,
		Enumerable:   false,
		Configurable: true,
		Writable:     true,
	})
	return Value{Type: VTJSObject, Num: objID}
}

// jsIntlCreateNumberFormat allocates one Intl.NumberFormat instance with normalized locale state.
func (vm *VM) jsIntlCreateNumberFormat(args []Value) Value {
	profile, localeTag := vm.jsIntlResolveLocaleProfile(jsArgOrUndefined(args, 0))
	options := Value{Type: VTJSUndefined}
	if len(args) > 1 {
		options = args[1]
	}
	style := strings.ToLower(strings.TrimSpace(vm.jsIntlOptionString(options, "style")))
	if style == "" {
		style = "decimal"
	}
	useGrouping := true
	if v, ok := vm.jsIntlOptionBool(options, "useGrouping"); ok {
		useGrouping = v
	}
	digits := vm.jsIntlResolveFractionDigits(style, options)
	currencyCode := strings.TrimSpace(vm.jsIntlOptionString(options, "currency"))
	currencySymbol, currencySpacing := jsIntlCurrencySymbol(currencyCode, profile)
	objID := vm.allocJSID()
	obj := make(map[string]Value, 9)
	obj["__js_type"] = NewString("Intl.NumberFormat")
	obj["__js_ctor"] = NewString("IntlNumberFormat")
	obj["__js_intl_locale"] = NewString(localeTag)
	obj["__js_intl_style"] = NewString(style)
	obj["__js_intl_use_grouping"] = NewBool(useGrouping)
	obj["__js_intl_digits"] = NewInteger(int64(digits))
	obj["__js_intl_currency_symbol"] = NewString(currencySymbol)
	obj["__js_intl_currency_spacing"] = NewString(currencySpacing)
	obj["format"] = vm.jsIntlCreateMethodFunction(objID, "format", "IntlNumberFormatFormat")
	obj["formatToParts"] = vm.jsIntlCreateMethodFunction(objID, "formatToParts", "IntlNumberFormatFormatToParts")
	vm.jsObjectItems[objID] = obj
	vm.jsIntlNumberFormatItems[objID] = &jsIntlNumberFormatObject{
		localeTag:       localeTag,
		style:           style,
		digits:          digits,
		useGrouping:     useGrouping,
		currencySymbol:  currencySymbol,
		currencySpacing: currencySpacing,
	}
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 7)
	vm.jsSetDescriptor(objID, "format", jsPropertyDescriptor{
		Value:        obj["format"],
		HasValue:     true,
		Enumerable:   false,
		Configurable: true,
		Writable:     true,
	})
	vm.jsSetDescriptor(objID, "formatToParts", jsPropertyDescriptor{
		Value:        obj["formatToParts"],
		HasValue:     true,
		Enumerable:   false,
		Configurable: true,
		Writable:     true,
	})
	return Value{Type: VTJSObject, Num: objID}
}

// jsIntlCreateCollator allocates one Intl.Collator instance with normalized locale state.
func (vm *VM) jsIntlCreateCollator(args []Value) Value {
	_, localeTag := vm.jsIntlResolveLocaleProfile(jsArgOrUndefined(args, 0))
	options := Value{Type: VTJSUndefined}
	if len(args) > 1 {
		options = args[1]
	}

	usage := strings.ToLower(strings.TrimSpace(vm.jsIntlOptionString(options, "usage")))
	if usage != "search" {
		usage = "sort"
	}

	sensitivity := strings.ToLower(strings.TrimSpace(vm.jsIntlOptionString(options, "sensitivity")))
	if sensitivity == "" {
		if usage == "sort" {
			sensitivity = "variant"
		} else {
			sensitivity = "base"
		}
	}

	ignorePunct, _ := vm.jsIntlOptionBool(options, "ignorePunctuation")
	numeric, _ := vm.jsIntlOptionBool(options, "numeric")
	caseFirst := strings.ToLower(strings.TrimSpace(vm.jsIntlOptionString(options, "caseFirst")))
	if caseFirst != "upper" && caseFirst != "lower" && caseFirst != "false" {
		caseFirst = "false"
	}

	objID := vm.allocJSID()
	obj := make(map[string]Value, 10)
	obj["__js_type"] = NewString("Intl.Collator")
	obj["__js_ctor"] = NewString("IntlCollator")
	obj["__js_intl_locale"] = NewString(localeTag)
	obj["__js_intl_usage"] = NewString(usage)
	obj["__js_intl_sensitivity"] = NewString(sensitivity)
	obj["__js_intl_ignore_punct"] = NewBool(ignorePunct)
	obj["__js_intl_numeric"] = NewBool(numeric)
	obj["__js_intl_case_first"] = NewString(caseFirst)
	obj["compare"] = vm.jsIntlCreateMethodFunction(objID, "compare", "IntlCollatorCompare")

	vm.jsObjectItems[objID] = obj
	vm.jsIntlCollatorItems[objID] = &jsIntlCollatorObject{
		localeTag:   localeTag,
		usage:       usage,
		sensitivity: sensitivity,
		ignorePunct: ignorePunct,
		numeric:     numeric,
		caseFirst:   caseFirst,
	}
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 6)
	vm.jsSetDescriptor(objID, "compare", jsPropertyDescriptor{
		Value:        obj["compare"],
		HasValue:     true,
		Enumerable:   false,
		Configurable: true,
		Writable:     true,
	})
	return Value{Type: VTJSObject, Num: objID}
}

// jsIntlCreatePluralRules allocates one Intl.PluralRules instance with normalized locale state.
func (vm *VM) jsIntlCreatePluralRules(args []Value) Value {
	_, localeTag := vm.jsIntlResolveLocaleProfile(jsArgOrUndefined(args, 0))
	options := Value{Type: VTJSUndefined}
	if len(args) > 1 {
		options = args[1]
	}

	style := strings.ToLower(strings.TrimSpace(vm.jsIntlOptionString(options, "type")))
	if style != "ordinal" {
		style = "cardinal"
	}

	objID := vm.allocJSID()
	obj := make(map[string]Value, 6)
	obj["__js_type"] = NewString("Intl.PluralRules")
	obj["__js_ctor"] = NewString("IntlPluralRules")
	obj["__js_intl_locale"] = NewString(localeTag)
	obj["__js_intl_style"] = NewString(style)
	obj["select"] = vm.jsIntlCreateMethodFunction(objID, "select", "IntlPluralRulesSelect")

	vm.jsObjectItems[objID] = obj
	vm.jsIntlPluralRulesItems[objID] = &jsIntlPluralRulesObject{
		localeTag: localeTag,
		style:     style,
	}
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 6)
	vm.jsSetDescriptor(objID, "select", jsPropertyDescriptor{
		Value:        obj["select"],
		HasValue:     true,
		Enumerable:   false,
		Configurable: true,
		Writable:     true,
	})
	return Value{Type: VTJSObject, Num: objID}
}

// jsIntlCreateRelativeTimeFormat allocates one Intl.RelativeTimeFormat instance with normalized locale state.
func (vm *VM) jsIntlCreateRelativeTimeFormat(args []Value) Value {
	_, localeTag := vm.jsIntlResolveLocaleProfile(jsArgOrUndefined(args, 0))
	options := Value{Type: VTJSUndefined}
	if len(args) > 1 {
		options = args[1]
	}

	numeric := strings.ToLower(strings.TrimSpace(vm.jsIntlOptionString(options, "numeric")))
	if numeric != "auto" {
		numeric = "always"
	}

	style := strings.ToLower(strings.TrimSpace(vm.jsIntlOptionString(options, "style")))
	if style != "short" && style != "narrow" {
		style = "long"
	}

	objID := vm.allocJSID()
	obj := make(map[string]Value, 8)
	obj["__js_type"] = NewString("Intl.RelativeTimeFormat")
	obj["__js_ctor"] = NewString("IntlRelativeTimeFormat")
	obj["__js_intl_locale"] = NewString(localeTag)
	obj["__js_intl_numeric"] = NewString(numeric)
	obj["__js_intl_style"] = NewString(style)
	obj["format"] = vm.jsIntlCreateMethodFunction(objID, "format", "IntlRelativeTimeFormatFormat")
	obj["formatToParts"] = vm.jsIntlCreateMethodFunction(objID, "formatToParts", "IntlRelativeTimeFormatFormatToParts")

	vm.jsObjectItems[objID] = obj
	vm.jsIntlRelativeTimeFormatItems[objID] = &jsIntlRelativeTimeFormatObject{
		localeTag: localeTag,
		numeric:   numeric,
		style:     style,
	}
	vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 7)
	vm.jsSetDescriptor(objID, "format", jsPropertyDescriptor{
		Value:        obj["format"],
		HasValue:     true,
		Enumerable:   false,
		Configurable: true,
		Writable:     true,
	})
	vm.jsSetDescriptor(objID, "formatToParts", jsPropertyDescriptor{
		Value:        obj["formatToParts"],
		HasValue:     true,
		Enumerable:   false,
		Configurable: true,
		Writable:     true,
	})
	return Value{Type: VTJSObject, Num: objID}
}

func (vm *VM) jsIntlRelativeTimeFormatFormat(callee Value, thisVal Value, args []Value) Value {
	ownerID, ok := vm.jsIntlMethodOwnerID(callee)
	if !ok {
		ownerID = thisVal.Num
	}
	obj, ok := vm.jsObjectItems[ownerID]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	inst := vm.jsIntlRelativeTimeFormatItems[ownerID]

	localeTag := obj["__js_intl_locale"].String()
	numeric := obj["__js_intl_numeric"].String()
	style := obj["__js_intl_style"].String()

	if inst != nil {
		localeTag = inst.localeTag
		numeric = inst.numeric
		style = inst.style
	}

	value := 0.0
	if len(args) > 0 {
		value = vm.jsToNumber(args[0]).Flt
	}
	unit := ""
	if len(args) > 1 {
		unit = strings.ToLower(strings.TrimSpace(args[1].String()))
	}

	result := vm.jsIntlPerformRelativeTimeFormat(value, unit, localeTag, numeric, style)
	return NewString(result)
}

// jsIntlRelativeTimeFormatFormatToParts decomposes one relative time into its constituent formatting parts.
func (vm *VM) jsIntlRelativeTimeFormatFormatToParts(callee Value, thisVal Value, args []Value) Value {
	ownerID, ok := vm.jsIntlMethodOwnerID(callee)
	if !ok {
		ownerID = thisVal.Num
	}
	obj, ok := vm.jsObjectItems[ownerID]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	inst := vm.jsIntlRelativeTimeFormatItems[ownerID]

	localeTag := obj["__js_intl_locale"].String()
	numeric := obj["__js_intl_numeric"].String()
	style := obj["__js_intl_style"].String()

	if inst != nil {
		localeTag = inst.localeTag
		numeric = inst.numeric
		style = inst.style
	}

	value := 0.0
	if len(args) > 0 {
		value = vm.jsToNumber(args[0]).Flt
	}
	unit := ""
	if len(args) > 1 {
		unit = strings.ToLower(strings.TrimSpace(args[1].String()))
	}

	parts := vm.jsIntlDecomposeRelativeTime(value, unit, localeTag, numeric, style)
	return vm.jsCreateIntlPartArray(parts)
}

// jsIntlPerformRelativeTimeFormat renders a relative time string.
func (vm *VM) jsIntlPerformRelativeTimeFormat(value float64, unit, localeTag, numeric, style string) string {
	parts := vm.jsIntlDecomposeRelativeTime(value, unit, localeTag, numeric, style)
	var sb strings.Builder
	for _, p := range parts {
		sb.WriteString(p.Value)
	}
	return sb.String()
}

// jsIntlDecomposeRelativeTime breaks down a relative time into segments.
func (vm *VM) jsIntlDecomposeRelativeTime(value float64, unit, localeTag, numeric, style string) []jsIntlPart {
	lang := strings.Split(localeTag, "-")[0]
	isFuture := value >= 0
	absVal := math.Abs(value)
	unit = strings.TrimSuffix(unit, "s") // normalize to singular

	// Try numeric: "auto" first
	if numeric == "auto" {
		if absVal == 0 {
			switch lang {
			case "en":
				return []jsIntlPart{{Type: "literal", Value: "now"}}
			case "pt":
				return []jsIntlPart{{Type: "literal", Value: "agora"}}
			}
		} else if absVal == 1 {
			if unit == "day" {
				if !isFuture {
					switch lang {
					case "en":
						return []jsIntlPart{{Type: "literal", Value: "yesterday"}}
					case "pt":
						return []jsIntlPart{{Type: "literal", Value: "ontem"}}
					}
				} else {
					switch lang {
					case "en":
						return []jsIntlPart{{Type: "literal", Value: "tomorrow"}}
					case "pt":
						return []jsIntlPart{{Type: "literal", Value: "amanhã"}}
					}
				}
			}
		}
	}

	// Default fallback (numeric: "always" or no special "auto" match)
	profile := builtinLocaleProfileForTag(localeTag)
	valStr := localizedNumberString(absVal, 0, profile, true)

	var res []jsIntlPart
	if !isFuture {
		switch lang {
		case "en":
			res = append(res, jsIntlPart{Type: "integer", Value: valStr})
			res = append(res, jsIntlPart{Type: "literal", Value: " " + unit})
			if absVal != 1 {
				res[1].Value += "s"
			}
			res = append(res, jsIntlPart{Type: "literal", Value: " ago"})
		case "pt":
			res = append(res, jsIntlPart{Type: "literal", Value: "há "})
			res = append(res, jsIntlPart{Type: "integer", Value: valStr})
			res = append(res, jsIntlPart{Type: "literal", Value: " " + unit})
			if absVal != 1 {
				res[2].Value += "s"
			}
		default:
			res = append(res, jsIntlPart{Type: "integer", Value: valStr})
			res = append(res, jsIntlPart{Type: "literal", Value: " " + unit + " ago"})
		}
	} else {
		switch lang {
		case "en":
			res = append(res, jsIntlPart{Type: "literal", Value: "in "})
			res = append(res, jsIntlPart{Type: "integer", Value: valStr})
			res = append(res, jsIntlPart{Type: "literal", Value: " " + unit})
			if absVal != 1 {
				res[2].Value += "s"
			}
		case "pt":
			res = append(res, jsIntlPart{Type: "literal", Value: "em "})
			res = append(res, jsIntlPart{Type: "integer", Value: valStr})
			res = append(res, jsIntlPart{Type: "literal", Value: " " + unit})
			if absVal != 1 {
				res[2].Value += "s"
			}
		default:
			res = append(res, jsIntlPart{Type: "literal", Value: "in "})
			res = append(res, jsIntlPart{Type: "integer", Value: valStr})
			res = append(res, jsIntlPart{Type: "literal", Value: " " + unit})
		}
	}

	return res
}

// jsIntlPluralRulesSelect determines the plural category of a number.
func (vm *VM) jsIntlPluralRulesSelect(callee Value, thisVal Value, args []Value) Value {
	ownerID, ok := vm.jsIntlMethodOwnerID(callee)
	if !ok {
		ownerID = thisVal.Num
	}
	obj, ok := vm.jsObjectItems[ownerID]
	if !ok {
		return NewString("other")
	}
	inst := vm.jsIntlPluralRulesItems[ownerID]

	localeTag := obj["__js_intl_locale"].String()
	style := obj["__js_intl_style"].String()

	if inst != nil {
		localeTag = inst.localeTag
		style = inst.style
	}

	n := 0.0
	if len(args) > 0 {
		n = vm.jsToNumber(args[0]).Flt
	}

	category := vm.jsIntlPerformPluralSelect(n, localeTag, style)
	return NewString(category)
}

// jsIntlPerformPluralSelect determines the plural category using CLDR-like rules.
func (vm *VM) jsIntlPerformPluralSelect(n float64, localeTag, style string) string {
	lang := strings.Split(localeTag, "-")[0]

	if style == "ordinal" {
		switch lang {
		case "en":
			v := int(math.Abs(n)) % 100
			if v == 11 || v == 12 || v == 13 {
				return "other"
			}
			switch int(math.Abs(n)) % 10 {
			case 1:
				return "one"
			case 2:
				return "two"
			case 3:
				return "few"
			default:
				return "other"
			}
		}
		// Default ordinal is other
		return "other"
	}

	// Default cardinal rules
	switch lang {
	case "en", "de", "nl", "sv", "da", "no", "nb", "nn", "et", "fi", "el", "he", "hu", "it", "es", "pt":
		if n == 1 {
			return "one"
		}
		return "other"
	case "fr", "br":
		if n >= 0 && n < 2 {
			return "one"
		}
		return "other"
	case "pl":
		v := int(math.Abs(n))
		if v == 1 {
			return "one"
		}
		if v%10 >= 2 && v%10 <= 4 && (v%100 < 12 || v%100 > 14) {
			return "few"
		}
		if v%10 == 0 || (v%10 >= 5 && v%10 <= 9) || (v%100 >= 11 && v%100 <= 14) {
			return "many"
		}
		return "other"
	case "ru", "uk", "be":
		v := int(math.Abs(n))
		if v%10 == 1 && v%100 != 11 {
			return "one"
		}
		if v%10 >= 2 && v%10 <= 4 && (v%100 < 12 || v%100 > 14) {
			return "few"
		}
		if v%10 == 0 || (v%10 >= 5 && v%10 <= 9) || (v%100 >= 11 && v%100 <= 14) {
			return "many"
		}
		return "other"
	}

	if n == 1 {
		return "one"
	}
	return "other"
}

// jsIntlCollatorCompare compares two strings using the collator state.
func (vm *VM) jsIntlCollatorCompare(callee Value, thisVal Value, args []Value) Value {
	ownerID, ok := vm.jsIntlMethodOwnerID(callee)
	if !ok {
		ownerID = thisVal.Num
	}
	obj, ok := vm.jsObjectItems[ownerID]
	if !ok {
		return NewInteger(0)
	}
	inst := vm.jsIntlCollatorItems[ownerID]

	// Default values if inst is nil (fallback to obj props)
	localeTag := obj["__js_intl_locale"].String()
	sensitivity := obj["__js_intl_sensitivity"].String()
	numeric := obj["__js_intl_numeric"].Num != 0
	caseFirst := obj["__js_intl_case_first"].String()
	ignorePunct := obj["__js_intl_ignore_punct"].Num != 0

	if inst != nil {
		localeTag = inst.localeTag
		sensitivity = inst.sensitivity
		numeric = inst.numeric
		caseFirst = inst.caseFirst
		ignorePunct = inst.ignorePunct
	}

	s1 := ""
	if len(args) > 0 {
		s1 = args[0].String()
	}
	s2 := ""
	if len(args) > 1 {
		s2 = args[1].String()
	}

	result := vm.jsIntlPerformCompare(s1, s2, localeTag, sensitivity, numeric, caseFirst, ignorePunct)
	return NewInteger(int64(result))
}

// jsIntlPerformCompare executes the locale-aware comparison logic.
func (vm *VM) jsIntlPerformCompare(s1, s2, localeTag, sensitivity string, numeric bool, caseFirst string, ignorePunct bool) int {
	if s1 == s2 {
		return 0
	}

	if ignorePunct {
		s1 = jsIntlStripPunct(s1)
		s2 = jsIntlStripPunct(s2)
	}

	// For basic implementation, we support sensitivity:
	// "base": only base letters (a == A == á)
	// "accent": base + accents (a == A, a != á)
	// "case": base + case (a == á, a != A)
	// "variant": all differences (default)

	switch sensitivity {
	case "base":
		s1 = jsIntlFoldPrimary(s1)
		s2 = jsIntlFoldPrimary(s2)
	case "accent":
		s1 = jsIntlFoldAccent(s1)
		s2 = jsIntlFoldAccent(s2)
	case "case":
		s1 = jsIntlFoldCase(s1)
		s2 = jsIntlFoldCase(s2)
	}

	if numeric {
		return jsIntlCompareNumeric(s1, s2)
	}

	if caseFirst == "upper" {
		return jsIntlCompareCaseFirst(s1, s2, true)
	} else if caseFirst == "lower" {
		return jsIntlCompareCaseFirst(s1, s2, false)
	}

	return strings.Compare(s1, s2)
}

func jsIntlStripPunct(s string) string {
	var sb strings.Builder
	for _, r := range s {
		if !((r >= 33 && r <= 47) || (r >= 58 && r <= 64) || (r >= 91 && r <= 96) || (r >= 123 && r <= 126)) {
			sb.WriteRune(r)
		}
	}
	return sb.String()
}

func jsIntlFoldPrimary(s string) string {
	// Approximation of primary folding (lower case + removing common accents)
	s = strings.ToLower(s)
	// Basic accent removal for common Latin-1
	var sb strings.Builder
	for _, r := range s {
		switch r {
		case 'á', 'à', 'â', 'ã', 'ä', 'å':
			sb.WriteRune('a')
		case 'é', 'è', 'ê', 'ë':
			sb.WriteRune('e')
		case 'í', 'ì', 'î', 'ï':
			sb.WriteRune('i')
		case 'ó', 'ò', 'ô', 'õ', 'ö':
			sb.WriteRune('o')
		case 'ú', 'ù', 'û', 'ü':
			sb.WriteRune('u')
		case 'ç':
			sb.WriteRune('c')
		case 'ñ':
			sb.WriteRune('n')
		default:
			sb.WriteRune(r)
		}
	}
	return sb.String()
}

func jsIntlFoldAccent(s string) string {
	// Case-insensitive but accent-sensitive
	return strings.ToLower(s)
}

func jsIntlFoldCase(s string) string {
	// Accent-insensitive but case-sensitive
	// We'll just remove accents but keep case for now
	var sb strings.Builder
	for _, r := range s {
		switch r {
		case 'á', 'à', 'â', 'ã', 'ä', 'å':
			sb.WriteRune('a')
		case 'Á', 'À', 'Â', 'Ã', 'Ä', 'Å':
			sb.WriteRune('A')
		case 'é', 'è', 'ê', 'ë':
			sb.WriteRune('e')
		case 'É', 'È', 'Ê', 'Ë':
			sb.WriteRune('E')
		case 'í', 'ì', 'î', 'ï':
			sb.WriteRune('i')
		case 'Í', 'Ì', 'Î', 'Ï':
			sb.WriteRune('I')
		case 'ó', 'ò', 'ô', 'õ', 'ö':
			sb.WriteRune('o')
		case 'Ó', 'Ò', 'Ô', 'Õ', 'Ö':
			sb.WriteRune('O')
		case 'ú', 'ù', 'û', 'ü':
			sb.WriteRune('u')
		case 'Ú', 'Ù', 'Û', 'Ü':
			sb.WriteRune('U')
		case 'ç':
			sb.WriteRune('c')
		case 'Ç':
			sb.WriteRune('C')
		case 'ñ':
			sb.WriteRune('n')
		case 'Ñ':
			sb.WriteRune('N')
		default:
			sb.WriteRune(r)
		}
	}
	return sb.String()
}

func jsIntlCompareNumeric(s1, s2 string) int {
	i, j := 0, 0
	for i < len(s1) && j < len(s2) {
		c1, c2 := s1[i], s2[j]
		if c1 >= '0' && c1 <= '9' && c2 >= '0' && c2 <= '9' {
			// Parse first number
			startI := i
			for i < len(s1) && s1[i] >= '0' && s1[i] <= '9' {
				i++
			}
			n1, _ := strconv.ParseInt(s1[startI:i], 10, 64)

			// Parse second number
			startJ := j
			for j < len(s2) && s2[j] >= '0' && s2[j] <= '9' {
				j++
			}
			n2, _ := strconv.ParseInt(s2[startJ:j], 10, 64)

			if n1 < n2 {
				return -1
			}
			if n1 > n2 {
				return 1
			}
			// If equal, continue to next part
		} else {
			if c1 < c2 {
				return -1
			}
			if c1 > c2 {
				return 1
			}
			i++
			j++
		}
	}
	if i < len(s1) {
		return 1
	}
	if j < len(s2) {
		return -1
	}
	return 0
}

func jsIntlCompareCaseFirst(s1, s2 string, upperFirst bool) int {
	res := strings.Compare(strings.ToLower(s1), strings.ToLower(s2))
	if res != 0 {
		return res
	}
	if upperFirst {
		return strings.Compare(s1, s2) // ASCII: A < a
	}
	return strings.Compare(s2, s1) // Reversed: a < A
}

// jsIntlPart represents one segment of a formatted Intl result.
type jsIntlPart struct {
	Type  string
	Value string
}

// jsCreateIntlPartArray converts a slice of jsIntlPart into a JS Array of objects.
func (vm *VM) jsCreateIntlPartArray(parts []jsIntlPart) Value {
	values := make([]Value, len(parts))
	for i, part := range parts {
		objID := vm.allocJSID()
		obj := make(map[string]Value, 2)
		obj["type"] = NewString(part.Type)
		obj["value"] = NewString(part.Value)
		vm.jsObjectItems[objID] = obj
		vm.jsPropertyItems[objID] = make(map[string]jsPropertyDescriptor, 2)
		values[i] = Value{Type: VTJSObject, Num: objID}
	}
	return ValueFromValueSlice(values)
}

// jsIntlDateTimeFormatFormat renders one date value using the formatter state stored on the Intl instance.
func (vm *VM) jsIntlDateTimeFormatFormat(callee Value, thisVal Value, args []Value) Value {
	ownerID, ok := vm.jsIntlMethodOwnerID(callee)
	if !ok {
		ownerID = thisVal.Num
	}
	obj, ok := vm.jsObjectItems[ownerID]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	inst := vm.jsIntlDateTimeFormatItems[ownerID]
	localeTag := obj["__js_intl_locale"].String()
	layout := obj["__js_intl_layout"].String()
	if inst != nil {
		localeTag = inst.localeTag
		layout = inst.layout
	}
	profile := builtinLocaleProfileForTag(localeTag)
	value := vm.jsIntlResolveDateTimeValue(args)
	if strings.TrimSpace(layout) == "" {
		layout = profile.shortDateLayout + " " + profile.longTimeLayout
	}
	return NewString(localizedFormat(value, layout, profile))
}

// jsIntlDateTimeFormatFormatToParts decomposes one date value into its constituent formatting parts.
func (vm *VM) jsIntlDateTimeFormatFormatToParts(callee Value, thisVal Value, args []Value) Value {
	ownerID, ok := vm.jsIntlMethodOwnerID(callee)
	if !ok {
		ownerID = thisVal.Num
	}
	obj, ok := vm.jsObjectItems[ownerID]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	inst := vm.jsIntlDateTimeFormatItems[ownerID]
	localeTag := obj["__js_intl_locale"].String()
	layout := obj["__js_intl_layout"].String()
	if inst != nil {
		localeTag = inst.localeTag
		layout = inst.layout
	}
	profile := builtinLocaleProfileForTag(localeTag)
	value := vm.jsIntlResolveDateTimeValue(args)
	if strings.TrimSpace(layout) == "" {
		layout = profile.shortDateLayout + " " + profile.longTimeLayout
	}
	parts := vm.jsIntlDecomposeDateTime(value, layout, profile)
	return vm.jsCreateIntlPartArray(parts)
}

// jsIntlResolveDateTimeValue extracts a time.Time value from the formatter method arguments.
func (vm *VM) jsIntlResolveDateTimeValue(args []Value) time.Time {
	value := time.Now().In(time.UTC)
	if len(args) > 0 {
		candidate := resolveCallable(vm, args[0])
		if candidate.Type == VTJSUndefined || candidate.Type == VTNull {
			// use current time
		} else {
			parsed := valueToTimeInLocale(vm, candidate)
			if !parsed.IsZero() {
				value = parsed
			}
		}
	}
	return value
}

// jsIntlDecomposeDateTime parses a Go layout string and renders each part against a time value.
func (vm *VM) jsIntlDecomposeDateTime(value time.Time, layout string, profile builtinLocaleProfile) []jsIntlPart {
	tokens := []struct {
		Pattern string
		Type    string
	}{
		{"Monday", "weekday"}, {"Mon", "weekday"},
		{"January", "month"}, {"Jan", "month"},
		{"2006", "year"},
		{"MST", "timeZoneName"}, {"-0700", "timeZoneName"}, {"Z0700", "timeZoneName"},
		{"01", "month"}, {"02", "day"}, {"_2", "day"}, {"03", "hour"}, {"04", "minute"}, {"05", "second"}, {"06", "year"},
		{"15", "hour"}, {"PM", "dayPeriod"}, {"pm", "dayPeriod"},
		{"1", "month"}, {"2", "day"}, {"3", "hour"}, {"4", "minute"}, {"5", "second"},
	}

	var parts []jsIntlPart
	remaining := layout
	for len(remaining) > 0 {
		found := false
		for _, t := range tokens {
			if strings.HasPrefix(remaining, t.Pattern) {
				val := localizedFormat(value, t.Pattern, profile)
				parts = append(parts, jsIntlPart{Type: t.Type, Value: val})
				remaining = remaining[len(t.Pattern):]
				found = true
				break
			}
		}
		if !found {
			// Collect literal characters
			literal := string(remaining[0])
			if len(parts) > 0 && parts[len(parts)-1].Type == "literal" {
				parts[len(parts)-1].Value += literal
			} else {
				parts = append(parts, jsIntlPart{Type: "literal", Value: literal})
			}
			remaining = remaining[1:]
		}
	}
	return parts
}

// jsIntlNumberFormatFormat renders one number using the formatter state stored on the Intl instance.
func (vm *VM) jsIntlNumberFormatFormat(callee Value, thisVal Value, args []Value) Value {
	ownerID, ok := vm.jsIntlMethodOwnerID(callee)
	if !ok {
		ownerID = thisVal.Num
	}
	obj, ok := vm.jsObjectItems[ownerID]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	inst := vm.jsIntlNumberFormatItems[ownerID]
	localeTag := obj["__js_intl_locale"].String()
	style := strings.ToLower(obj["__js_intl_style"].String())
	digits := int(obj["__js_intl_digits"].Num)
	useGrouping := obj["__js_intl_use_grouping"].Num != 0
	currencySymbol := obj["__js_intl_currency_symbol"].String()
	currencySpacing := obj["__js_intl_currency_spacing"].String()
	if inst != nil {
		localeTag = inst.localeTag
		style = inst.style
		digits = inst.digits
		useGrouping = inst.useGrouping
		currencySymbol = inst.currencySymbol
		currencySpacing = inst.currencySpacing
	}
	profile := builtinLocaleProfileForTag(localeTag)
	value := 0.0
	if len(args) > 0 {
		value = vm.jsToNumber(args[0]).Flt
	}
	switch style {
	case "currency":
		formatted := localizedNumberString(value, digits, profile, useGrouping)
		if strings.TrimSpace(currencySymbol) == "" {
			currencySymbol = profile.currencySymbol
			currencySpacing = profile.currencySpacing
		}
		result := currencySymbol + currencySpacing + formatted
		if value < 0 {
			return NewString("-" + result)
		}
		return NewString(result)
	case "percent":
		formatted := localizedNumberString(value*100, digits, profile, useGrouping)
		return NewString(formatted + "%")
	default:
		return NewString(localizedNumberString(value, digits, profile, useGrouping))
	}
}

// jsIntlNumberFormatFormatToParts decomposes one number into its constituent formatting parts.
func (vm *VM) jsIntlNumberFormatFormatToParts(callee Value, thisVal Value, args []Value) Value {
	ownerID, ok := vm.jsIntlMethodOwnerID(callee)
	if !ok {
		ownerID = thisVal.Num
	}
	obj, ok := vm.jsObjectItems[ownerID]
	if !ok {
		return Value{Type: VTJSUndefined}
	}
	inst := vm.jsIntlNumberFormatItems[ownerID]
	localeTag := obj["__js_intl_locale"].String()
	style := strings.ToLower(obj["__js_intl_style"].String())
	digits := int(obj["__js_intl_digits"].Num)
	useGrouping := obj["__js_intl_use_grouping"].Num != 0
	currencySymbol := obj["__js_intl_currency_symbol"].String()
	currencySpacing := obj["__js_intl_currency_spacing"].String()
	if inst != nil {
		localeTag = inst.localeTag
		style = inst.style
		digits = inst.digits
		useGrouping = inst.useGrouping
		currencySymbol = inst.currencySymbol
		currencySpacing = inst.currencySpacing
	}
	profile := builtinLocaleProfileForTag(localeTag)
	value := 0.0
	if len(args) > 0 {
		value = vm.jsToNumber(args[0]).Flt
	}
	parts := vm.jsIntlDecomposeNumber(value, digits, profile, useGrouping, style, currencySymbol, currencySpacing)
	return vm.jsCreateIntlPartArray(parts)
}

// jsIntlDecomposeNumber breaks down a numeric value into segments based on locale and style.
func (vm *VM) jsIntlDecomposeNumber(value float64, digits int, profile builtinLocaleProfile, useGrouping bool, style string, currencySymbol string, currencySpacing string) []jsIntlPart {
	if digits < 0 {
		digits = 0
	}
	absVal := math.Abs(value)
	if style == "percent" {
		absVal *= 100
	}
	raw := strconv.FormatFloat(absVal, 'f', digits, 64)
	parts := strings.SplitN(raw, ".", 2)
	integerStr := parts[0]

	var res []jsIntlPart

	if value < 0 {
		res = append(res, jsIntlPart{Type: "minusSign", Value: "-"})
	}

	if style == "currency" {
		if strings.TrimSpace(currencySymbol) == "" {
			currencySymbol = profile.currencySymbol
			currencySpacing = profile.currencySpacing
		}
		res = append(res, jsIntlPart{Type: "currency", Value: currencySymbol})
		if currencySpacing != "" {
			res = append(res, jsIntlPart{Type: "literal", Value: currencySpacing})
		}
	}

	if useGrouping && len(integerStr) > 3 {
		leading := len(integerStr) % 3
		if leading == 0 {
			leading = 3
		}
		res = append(res, jsIntlPart{Type: "integer", Value: integerStr[:leading]})
		for i := leading; i < len(integerStr); i += 3 {
			res = append(res, jsIntlPart{Type: "group", Value: profile.thousandSeparator})
			res = append(res, jsIntlPart{Type: "integer", Value: integerStr[i : i+3]})
		}
	} else {
		res = append(res, jsIntlPart{Type: "integer", Value: integerStr})
	}

	if digits > 0 && len(parts) > 1 {
		res = append(res, jsIntlPart{Type: "decimal", Value: profile.decimalSeparator})
		res = append(res, jsIntlPart{Type: "fraction", Value: parts[1]})
	}

	if style == "percent" {
		res = append(res, jsIntlPart{Type: "percentSymbol", Value: "%"})
	}

	return res
}

// jsIntlResolveLocaleProfile resolves one locale profile from Intl constructor arguments.
func (vm *VM) jsIntlResolveLocaleProfile(locales Value) (builtinLocaleProfile, string) {
	defaultProfile := builtinLocaleProfileForVM(vm)
	tag := vm.jsIntlLocaleTagFromValue(locales)
	if strings.TrimSpace(tag) == "" {
		return defaultProfile, defaultProfile.tag
	}
	profile := builtinLocaleProfileForTag(tag)
	return profile, profile.tag
}

// jsIntlLocaleTagFromValue extracts the first usable locale tag from one Intl locales argument.
func (vm *VM) jsIntlLocaleTagFromValue(locales Value) string {
	switch locales.Type {
	case VTJSUndefined, VTNull, VTEmpty:
		return ""
	case VTString:
		return jsNormalizeLocaleTag(locales.Str)
	case VTArray:
		if locales.Arr == nil {
			return ""
		}
		for i := 0; i < len(locales.Arr.Values); i++ {
			tag := jsNormalizeLocaleTag(locales.Arr.Values[i].String())
			if tag != "" {
				return tag
			}
		}
		return ""
	case VTJSObject, VTJSFunction:
		if length, ok, deferred := vm.jsArrayLikeLength(locales); ok && !deferred {
			for i := range length {
				if v, exists := vm.jsArrayLikeGetIndex(locales, i); exists {
					tag := jsNormalizeLocaleTag(v.String())
					if tag != "" {
						return tag
					}
				}
			}
		}
		return jsNormalizeLocaleTag(locales.String())
	default:
		return jsNormalizeLocaleTag(locales.String())
	}
}

// jsIntlOptionString reads one string option from an Intl options object.
func (vm *VM) jsIntlOptionString(options Value, name string) string {
	if options.Type != VTJSObject && options.Type != VTJSFunction {
		return ""
	}
	value, deferred := vm.jsMemberGet(options, name)
	if deferred || value.Type == VTJSUndefined || value.Type == VTNull {
		return ""
	}
	return strings.TrimSpace(value.String())
}

// jsIntlOptionBool reads one boolean option from an Intl options object.
func (vm *VM) jsIntlOptionBool(options Value, name string) (bool, bool) {
	if options.Type != VTJSObject && options.Type != VTJSFunction {
		return false, false
	}
	value, deferred := vm.jsMemberGet(options, name)
	if deferred || value.Type == VTJSUndefined || value.Type == VTNull {
		return false, false
	}
	switch value.Type {
	case VTBool:
		return value.Num != 0, true
	case VTInteger, VTDouble:
		return vm.jsToNumber(value).Flt != 0, true
	default:
		return strings.TrimSpace(value.String()) != "", true
	}
}

// jsIntlResolveFractionDigits selects the formatter precision for Intl.NumberFormat.
func (vm *VM) jsIntlResolveFractionDigits(style string, options Value) int {
	defaultDigits := 2
	switch style {
	case "percent":
		defaultDigits = 0
	case "currency":
		defaultDigits = 2
	}
	if digits, ok := vm.jsIntlOptionInt(options, "maximumFractionDigits"); ok {
		return digits
	}
	if digits, ok := vm.jsIntlOptionInt(options, "minimumFractionDigits"); ok {
		return digits
	}
	if digits, ok := vm.jsIntlOptionInt(options, "fractionDigits"); ok {
		return digits
	}
	return defaultDigits
}

// jsIntlOptionInt reads one integer option from an Intl options object.
func (vm *VM) jsIntlOptionInt(options Value, name string) (int, bool) {
	if options.Type != VTJSObject && options.Type != VTJSFunction {
		return 0, false
	}
	value, deferred := vm.jsMemberGet(options, name)
	if deferred || value.Type == VTJSUndefined || value.Type == VTNull {
		return 0, false
	}
	return int(math.Round(vm.jsToNumber(value).Flt)), true
}

// jsIntlResolveDateTimeLayout builds one locale-aware layout string from Intl.DateTimeFormat options.
func (vm *VM) jsIntlResolveDateTimeLayout(profile builtinLocaleProfile, options Value) string {
	dateStyle := strings.ToLower(vm.jsIntlOptionString(options, "dateStyle"))
	timeStyle := strings.ToLower(vm.jsIntlOptionString(options, "timeStyle"))
	hasDateTokens := vm.jsIntlHasDateTimeToken(options, "year") || vm.jsIntlHasDateTimeToken(options, "month") || vm.jsIntlHasDateTimeToken(options, "day") || vm.jsIntlHasDateTimeToken(options, "weekday")
	hasTimeTokens := vm.jsIntlHasDateTimeToken(options, "hour") || vm.jsIntlHasDateTimeToken(options, "minute") || vm.jsIntlHasDateTimeToken(options, "second")
	hour12, _ := vm.jsIntlOptionBool(options, "hour12")

	dateLayout := ""
	switch dateStyle {
	case "full", "long":
		dateLayout = localizedLongDateLayout(profile)
	case "medium", "short":
		dateLayout = profile.shortDateLayout
	}
	if dateLayout == "" && hasDateTokens {
		dateLayout = profile.shortDateLayout
	}

	timeLayout := ""
	switch timeStyle {
	case "full", "long":
		timeLayout = profile.longTimeLayout
	case "medium", "short":
		timeLayout = profile.shortTimeLayout
	}
	if timeLayout == "" && hasTimeTokens {
		if hour12 && strings.Contains(profile.shortTimeLayout, "PM") {
			timeLayout = profile.shortTimeLayout
		} else if hour12 && profile.tag != "en-US" {
			timeLayout = profile.shortTimeLayout
		} else {
			timeLayout = profile.longTimeLayout
		}
	}

	if dateLayout != "" && timeLayout != "" {
		return dateLayout + " " + timeLayout
	}
	if dateLayout != "" {
		return dateLayout
	}
	if timeLayout != "" {
		return timeLayout
	}
	return profile.shortDateLayout + " " + profile.longTimeLayout
}

// jsIntlHasDateTimeToken reports whether one Intl.DateTimeFormat option is present.
func (vm *VM) jsIntlHasDateTimeToken(options Value, name string) bool {
	if options.Type != VTJSObject && options.Type != VTJSFunction {
		return false
	}
	value, deferred := vm.jsMemberGet(options, name)
	return !deferred && value.Type != VTJSUndefined && value.Type != VTNull && strings.TrimSpace(value.String()) != ""
}

// jsIntlMethodOwnerID resolves the hidden owner object ID for one Intl formatter method.
func (vm *VM) jsIntlMethodOwnerID(callee Value) (int64, bool) {
	if callee.Type != VTJSFunction {
		return 0, false
	}
	obj, ok := vm.jsObjectItems[callee.Num]
	if !ok {
		return 0, false
	}
	owner, ok := obj["__js_intl_owner"]
	if !ok || owner.Type != VTInteger {
		return 0, false
	}
	return owner.Num, true
}

// jsNormalizeLocaleTag normalizes one locale string for matcher lookup.
func jsNormalizeLocaleTag(tag string) string {
	tag = strings.TrimSpace(tag)
	if tag == "" {
		return ""
	}
	tag = strings.ReplaceAll(tag, "_", "-")
	return tag
}

// jsIntlCurrencySymbol resolves the currency symbol and spacing for one Intl.NumberFormat currency
// code. It delegates to builtinCurrencySymbolForCode so that builtinLocaleProfiles in
// locale_format.go is the single source of truth shared by both VBScript and JScript.
func jsIntlCurrencySymbol(code string, profile builtinLocaleProfile) (string, string) {
	if strings.TrimSpace(code) == "" {
		return profile.currencySymbol, profile.currencySpacing
	}
	if symbol, spacing, ok := builtinCurrencySymbolForCode(code); ok {
		return symbol, spacing
	}
	// Unknown currency code: use the code itself as the display symbol.
	return strings.ToUpper(strings.TrimSpace(code)), " "
}
