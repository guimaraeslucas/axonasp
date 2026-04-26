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
	"strconv"
	"strings"
	"time"

	"github.com/goodsign/monday"
	"golang.org/x/text/language"
)

// builtinLocaleProfile stores date, time, and currency rendering rules for one locale.
type builtinLocaleProfile struct {
	tag               string
	mondayLocale      monday.Locale
	shortDateLayout   string
	longTimeLayout    string
	shortTimeLayout   string
	decimalSeparator  string
	thousandSeparator string
	currencySymbol    string
	currencySpacing   string
}

var builtinLocaleProfiles = []builtinLocaleProfile{
	{tag: "en-US", mondayLocale: monday.LocaleEnUS, shortDateLayout: "1/2/2006", longTimeLayout: "3:04:05 PM", shortTimeLayout: "3:04 PM", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "$", currencySpacing: ""},
	{tag: "en-GB", mondayLocale: monday.LocaleEnGB, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "£", currencySpacing: ""},
	{tag: "pt-BR", mondayLocale: monday.LocalePtBR, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "R$", currencySpacing: " "},
	{tag: "pt-PT", mondayLocale: monday.LocalePtPT, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "€", currencySpacing: " "},
	{tag: "es-ES", mondayLocale: monday.LocaleEsES, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "€", currencySpacing: " "},
	{tag: "fr-FR", mondayLocale: monday.LocaleFrFR, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: " ", currencySymbol: "€", currencySpacing: " "},
	{tag: "fr-CA", mondayLocale: monday.LocaleFrCA, shortDateLayout: "2006-01-02", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: " ", currencySymbol: "$", currencySpacing: ""},
	{tag: "de-DE", mondayLocale: monday.LocaleDeDE, shortDateLayout: "02.01.2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "€", currencySpacing: " "},
	{tag: "it-IT", mondayLocale: monday.LocaleItIT, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "€", currencySpacing: " "},
	{tag: "nl-NL", mondayLocale: monday.LocaleNlNL, shortDateLayout: "02-01-2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "€", currencySpacing: " "},
	{tag: "da-DK", mondayLocale: monday.LocaleDaDK, shortDateLayout: "02/01/2006", longTimeLayout: "15.04.05", shortTimeLayout: "15.04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "kr", currencySpacing: " "},
	{tag: "fi-FI", mondayLocale: monday.LocaleFiFI, shortDateLayout: "02.1.2006", longTimeLayout: "15.04.05", shortTimeLayout: "15.04", decimalSeparator: ",", thousandSeparator: " ", currencySymbol: "€", currencySpacing: " "},
	{tag: "nb-NO", mondayLocale: monday.LocaleNbNO, shortDateLayout: "02.01.2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: " ", currencySymbol: "kr", currencySpacing: " "},
	{tag: "pl-PL", mondayLocale: monday.LocalePlPL, shortDateLayout: "02.01.2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: " ", currencySymbol: "zł", currencySpacing: " "},
	{tag: "cs-CZ", mondayLocale: monday.LocaleCsCZ, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: " ", currencySymbol: "Kč", currencySpacing: " "},
	{tag: "bg-BG", mondayLocale: monday.LocaleBgBG, shortDateLayout: "2.01.2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: " ", currencySymbol: "лв", currencySpacing: " "},
	{tag: "ru-RU", mondayLocale: monday.LocaleRuRU, shortDateLayout: "02.01.2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: " ", currencySymbol: "₽", currencySpacing: " "},
	{tag: "uk-UA", mondayLocale: monday.LocaleUkUA, shortDateLayout: "02.01.2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: " ", currencySymbol: "₴", currencySpacing: " "},
	{tag: "zh-CN", mondayLocale: monday.LocaleZhCN, shortDateLayout: "2006/1/2", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "¥", currencySpacing: ""},
	{tag: "zh-TW", mondayLocale: monday.LocaleZhTW, shortDateLayout: "2006/1/2", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "NT$", currencySpacing: ""},
	{tag: "zh-HK", mondayLocale: monday.LocaleZhHK, shortDateLayout: "2006/1/2", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "HK$", currencySpacing: ""},
	{tag: "ko-KR", mondayLocale: monday.LocaleKoKR, shortDateLayout: "2006/1/2", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "₩", currencySpacing: ""},
	{tag: "ja-JP", mondayLocale: monday.LocaleJaJP, shortDateLayout: "2006/1/2", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "¥", currencySpacing: ""},
	{tag: "el-GR", mondayLocale: monday.LocaleElGR, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "€", currencySpacing: " "},
	{tag: "id-ID", mondayLocale: monday.LocaleIdID, shortDateLayout: "2/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "Rp", currencySpacing: " "},
	{tag: "tr-TR", mondayLocale: monday.LocaleTrTR, shortDateLayout: "2.01.2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "₺", currencySpacing: ""},
	{tag: "hr-HR", mondayLocale: monday.LocaleHrHR, shortDateLayout: "2.1.2006.", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "€", currencySpacing: " "},
	{tag: "sk-SK", mondayLocale: monday.LocaleSkSK, shortDateLayout: "2. 1. 2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: " ", currencySymbol: "€", currencySpacing: " "},
	{tag: "th-TH", mondayLocale: monday.LocaleThTH, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "฿", currencySpacing: ""},
	// English variants
	{tag: "en-AU", mondayLocale: monday.LocaleEnGB, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "$", currencySpacing: ""},
	{tag: "en-CA", mondayLocale: monday.LocaleEnUS, shortDateLayout: "1/2/2006", longTimeLayout: "3:04:05 PM", shortTimeLayout: "3:04 PM", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "$", currencySpacing: ""},
	{tag: "en-IN", mondayLocale: monday.LocaleEnUS, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "₹", currencySpacing: ""},
	{tag: "en-IE", mondayLocale: monday.LocaleEnGB, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "€", currencySpacing: " "},
	{tag: "en-NZ", mondayLocale: monday.LocaleEnGB, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "$", currencySpacing: ""},
	{tag: "en-ZA", mondayLocale: monday.LocaleEnGB, shortDateLayout: "2006/01/02", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "R", currencySpacing: " "},
	// Spanish variants
	{tag: "es-MX", mondayLocale: monday.LocaleEsES, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "$", currencySpacing: ""},
	{tag: "es-AR", mondayLocale: monday.LocaleEsES, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "$", currencySpacing: ""},
	{tag: "es-CO", mondayLocale: monday.LocaleEsES, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "$", currencySpacing: ""},
	{tag: "es-CL", mondayLocale: monday.LocaleEsES, shortDateLayout: "02-01-2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "$", currencySpacing: ""},
	{tag: "es-PE", mondayLocale: monday.LocaleEsES, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: ",", currencySymbol: "S/", currencySpacing: ""},
	// French variants
	{tag: "fr-BE", mondayLocale: monday.LocaleFrFR, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: " ", currencySymbol: "€", currencySpacing: " "},
	{tag: "fr-CH", mondayLocale: monday.LocaleFrFR, shortDateLayout: "02.01.2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: "'", currencySymbol: "CHF", currencySpacing: " "},
	// German variants
	{tag: "de-AT", mondayLocale: monday.LocaleDeDE, shortDateLayout: "02.01.2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "€", currencySpacing: " "},
	{tag: "de-CH", mondayLocale: monday.LocaleDeDE, shortDateLayout: "02.01.2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: "'", currencySymbol: "CHF", currencySpacing: " "},
	// Italian variant
	{tag: "it-CH", mondayLocale: monday.LocaleItIT, shortDateLayout: "02.01.2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ".", thousandSeparator: "'", currencySymbol: "CHF", currencySpacing: " "},
	// Dutch variant
	{tag: "nl-BE", mondayLocale: monday.LocaleNlNL, shortDateLayout: "02/01/2006", longTimeLayout: "15:04:05", shortTimeLayout: "15:04", decimalSeparator: ",", thousandSeparator: ".", currencySymbol: "€", currencySpacing: " "},
}

var builtinLocaleMatcher = newBuiltinLocaleMatcher()

// newBuiltinLocaleMatcher builds the language matcher used to map LCIDs to supported monday locales.
func newBuiltinLocaleMatcher() language.Matcher {
	supportedTags := make([]language.Tag, 0, len(builtinLocaleProfiles))
	for _, profile := range builtinLocaleProfiles {
		supportedTags = append(supportedTags, language.Make(profile.tag))
	}
	return language.NewMatcher(supportedTags)
}

// builtinLocaleProfileForVM resolves the current locale profile for one VM execution context.
func builtinLocaleProfileForVM(vm *VM) builtinLocaleProfile {
	return builtinLocaleProfileForLCID(builtinCurrentLCID(vm))
}

// builtinLocaleProfileForLCID resolves the closest supported locale profile for one LCID.
func builtinLocaleProfileForLCID(lcid int) builtinLocaleProfile {
	profile := builtinLocaleProfileForTag(GetGoLanguageFromMSLCID(MSLCID(lcid)))

	switch MSLCID(lcid) {
	case EnglishAustralia, EnglishCanada, EnglishNZ:
		profile.currencySymbol = "$"
		profile.currencySpacing = ""
	case EnglishIndia:
		profile.currencySymbol = "₹"
		profile.currencySpacing = ""
	case EnglishIreland:
		profile.currencySymbol = "€"
		profile.currencySpacing = " "
	case EnglishSouthAfr, AfrikaansSouthAfr:
		profile.currencySymbol = "R"
		profile.currencySpacing = " "
	case SpanishMexico, SpanishArgentina, SpanishColombia, SpanishChile:
		profile.currencySymbol = "$"
		profile.currencySpacing = ""
	case SpanishPeru:
		profile.currencySymbol = "S/"
		profile.currencySpacing = ""
	case FrenchSwitzerland, GermanSwitzerland, ItalianSwitzerland:
		profile.currencySymbol = "CHF"
		profile.currencySpacing = " "
	case DutchBelgium, FrenchBelgium, GermanAustria, ItalianItaly, PortuguesePortugal, SpanishSpain, FrenchFrance, GermanGermany, DutchNetherlands, GreekGreece, FinnishFinland, SlovakSlovakia, Croatian, Bulgarian:
		profile.currencySymbol = "€"
		profile.currencySpacing = " "
	case ChineseTaiwan:
		profile.currencySymbol = "NT$"
	case ChineseHongKong:
		profile.currencySymbol = "HK$"
	case JapaneseJapan, ChineseChina:
		profile.currencySymbol = "¥"
	case KoreanKorea:
		profile.currencySymbol = "₩"
	}

	return profile
}

// builtinLocaleProfileForTag resolves the closest supported locale profile for one BCP 47 language tag.
func builtinLocaleProfileForTag(tag string) builtinLocaleProfile {
	if strings.TrimSpace(tag) == "" {
		return builtinLocaleProfiles[0]
	}
	_, index, _ := builtinLocaleMatcher.Match(language.Make(tag))
	if index >= 0 && index < len(builtinLocaleProfiles) {
		return builtinLocaleProfiles[index]
	}
	return builtinLocaleProfiles[0]
}

// localizedMonthNames returns locale-aware month names for the selected language tag.
func localizedMonthNames(tag string, abbrev bool) []string {
	profile := builtinLocaleProfileForTag(tag)
	if abbrev {
		if names := monday.GetShortMonths(profile.mondayLocale); len(names) == 12 {
			return append([]string(nil), names...)
		}
	} else if names := monday.GetLongMonths(profile.mondayLocale); len(names) == 12 {
		return append([]string(nil), names...)
	}
	if abbrev {
		return []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}
	}
	return []string{"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}
}

// localizedWeekdayNames returns locale-aware weekday names for the selected language tag.
func localizedWeekdayNames(tag string, abbrev bool) []string {
	profile := builtinLocaleProfileForTag(tag)
	layout := "Monday"
	fallback := []string{"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}
	if abbrev {
		layout = "Mon"
		fallback = []string{"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"}
	}
	baseSunday := time.Date(2026, time.January, 4, 0, 0, 0, 0, time.UTC)
	names := make([]string, 0, 7)
	for offset := 0; offset < 7; offset++ {
		names = append(names, localizedFormat(baseSunday.AddDate(0, 0, offset), layout, profile))
	}
	for _, name := range names {
		if strings.TrimSpace(name) == "" {
			return fallback
		}
	}
	return names
}

// localizedFormat formats one time value using monday when the layout contains translated day or month names.
func localizedFormat(value time.Time, layout string, profile builtinLocaleProfile) string {
	return monday.Format(value, layout, profile.mondayLocale)
}

// localizedLongDateLayout returns the locale-specific long-date layout used by VBScript compatibility helpers.
func localizedLongDateLayout(profile builtinLocaleProfile) string {
	if layout, ok := monday.FullFormatsByLocale[profile.mondayLocale]; ok && strings.TrimSpace(layout) != "" {
		return layout
	}
	return "Monday, January 2, 2006"
}

// localizedDateTimeText renders one date/time value using the requested VBScript format constant.
func localizedDateTimeText(value time.Time, profile builtinLocaleProfile, formatType int) string {
	switch formatType {
	case 1:
		return localizedFormat(value, localizedLongDateLayout(profile), profile)
	case 2:
		return localizedFormat(value, profile.shortDateLayout, profile)
	case 3:
		return localizedFormat(value, profile.longTimeLayout, profile)
	case 4:
		return localizedFormat(value, profile.shortTimeLayout, profile)
	default:
		return localizedFormat(value, profile.shortDateLayout+" "+profile.longTimeLayout, profile)
	}
}

// localizedDateString renders implicit VTDate to string conversions using locale-aware date and time formats.
func localizedDateString(value time.Time, profile builtinLocaleProfile) string {
	if value.Hour() == 0 && value.Minute() == 0 && value.Second() == 0 {
		return localizedFormat(value, profile.shortDateLayout, profile)
	}
	return localizedFormat(value, profile.shortDateLayout+" "+profile.longTimeLayout, profile)
}

// localizedNumberString renders a floating-point number using locale-specific separators.
func localizedNumberString(value float64, digits int, profile builtinLocaleProfile, useGrouping bool) string {
	if digits < 0 {
		digits = 0
	}
	negative := value < 0
	if negative {
		value = -value
	}
	raw := strconv.FormatFloat(value, 'f', digits, 64)
	parts := strings.SplitN(raw, ".", 2)
	integerPart := parts[0]
	if useGrouping {
		integerPart = groupedIntegerString(integerPart, profile.thousandSeparator)
	}
	result := integerPart
	if digits > 0 {
		fractionPart := ""
		if len(parts) > 1 {
			fractionPart = parts[1]
		}
		result += profile.decimalSeparator + fractionPart
	}
	if negative {
		return "-" + result
	}
	return result
}

// groupedIntegerString injects thousand separators into one base-10 integer string.
func groupedIntegerString(value string, separator string) string {
	if separator == "" || len(value) <= 3 {
		return value
	}
	var builder strings.Builder
	leading := len(value) % 3
	if leading == 0 {
		leading = 3
	}
	builder.Grow(len(value) + (len(value)-1)/3*len(separator))
	builder.WriteString(value[:leading])
	for index := leading; index < len(value); index += 3 {
		builder.WriteString(separator)
		builder.WriteString(value[index : index+3])
	}
	return builder.String()
}

// localizedCurrencyString renders a currency value using locale-specific separators and symbol placement.
func localizedCurrencyString(value float64, digits int, profile builtinLocaleProfile) string {
	negative := value < 0
	if negative {
		value = -value
	}
	formattedNumber := localizedNumberString(value, digits, profile, true)
	result := profile.currencySymbol + profile.currencySpacing + formattedNumber
	if negative {
		return "-" + result
	}
	return result
}

// parseLocalizedTimeValue parses text using locale-priority date layouts and localized month/day names.
func parseLocalizedTimeValue(text string, location *time.Location, profile builtinLocaleProfile) time.Time {
	if strings.TrimSpace(text) == "" {
		return time.Time{}
	}
	layouts := []string{
		"2006-01-02 15:04:05",
		"2006-01-02 15:04",
		"2006-01-02",
		time.RFC3339,
		localizedLongDateLayout(profile),
		profile.shortDateLayout + " " + profile.longTimeLayout,
		profile.shortDateLayout + " " + profile.shortTimeLayout,
		profile.shortDateLayout,
		"15:04:05",
		"15:04",
		"3:04:05 PM",
		"3:04 PM",
		"01/02/2006 15:04:05",
		"01/02/2006 15:04",
		"01/02/2006",
		"02/01/2006 15:04:05",
		"02/01/2006 15:04",
		"02/01/2006",
	}
	seen := make(map[string]struct{}, len(layouts))
	for _, layout := range layouts {
		if strings.TrimSpace(layout) == "" {
			continue
		}
		if _, exists := seen[layout]; exists {
			continue
		}
		seen[layout] = struct{}{}
		if parsed, err := monday.ParseInLocation(layout, text, location, profile.mondayLocale); err == nil {
			return parsed
		}
		if parsed, err := time.ParseInLocation(layout, text, location); err == nil {
			return parsed
		}
	}
	return time.Time{}
}
