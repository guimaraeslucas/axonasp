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

// MSLCID represents Microsoft Locale IDs (Decimal)
type MSLCID int

const (
	// Inglês
	EnglishUS        MSLCID = 1033
	EnglishUK        MSLCID = 2057
	EnglishAustralia MSLCID = 3081
	EnglishCanada    MSLCID = 4105
	EnglishIndia     MSLCID = 16393
	EnglishIreland   MSLCID = 6153
	EnglishNZ        MSLCID = 5129
	EnglishSouthAfr  MSLCID = 7177

	// Espanhol
	SpanishSpain     MSLCID = 1034
	SpanishMexico    MSLCID = 2058
	SpanishArgentina MSLCID = 11274
	SpanishColombia  MSLCID = 9226
	SpanishChile     MSLCID = 13322
	SpanishPeru      MSLCID = 10250

	// Português
	PortugueseBrazil   MSLCID = 1046
	PortuguesePortugal MSLCID = 2070

	// Francês
	FrenchFrance      MSLCID = 1036
	FrenchCanada      MSLCID = 3084
	FrenchSwitzerland MSLCID = 4108
	FrenchBelgium     MSLCID = 2060

	// Alemão
	GermanGermany     MSLCID = 1031
	GermanAustria     MSLCID = 3079
	GermanSwitzerland MSLCID = 2055

	// Chinês
	ChineseChina    MSLCID = 2052
	ChineseTaiwan   MSLCID = 1028
	ChineseHongKong MSLCID = 3076

	// Outros Idiomas Asiáticos
	JapaneseJapan MSLCID = 1041
	KoreanKorea   MSLCID = 1042
	HindiIndia    MSLCID = 1081
	BengaliIndia  MSLCID = 1093
	TamilIndia    MSLCID = 1097
	Indonesian    MSLCID = 1057
	MalayMalaysia MSLCID = 1086
	Vietnamese    MSLCID = 1066
	ThaiThailand  MSLCID = 1054
	TagalogPhils  MSLCID = 1124

	// Árabe e Oriente Médio
	ArabicSaudiArab MSLCID = 1025
	ArabicEgypt     MSLCID = 3073
	ArabicUAE       MSLCID = 14337
	HebrewIsrael    MSLCID = 1037
	PersianIran     MSLCID = 1065
	UrduPakistan    MSLCID = 1056

	// Leste Europeu e Rússia
	RussianRussia    MSLCID = 1049
	Ukrainian        MSLCID = 1058
	PolishPoland     MSLCID = 1045
	CzechCzechRep    MSLCID = 1029
	SlovakSlovakia   MSLCID = 1051
	HungarianHungary MSLCID = 1038
	RomanianRomania  MSLCID = 1048
	Bulgarian        MSLCID = 1026
	Croatian         MSLCID = 1050

	// Outros Europeus (Ocidentais e Nórdicos)
	ItalianItaly       MSLCID = 1040
	ItalianSwitzerland MSLCID = 2064
	DutchNetherlands   MSLCID = 1043
	DutchBelgium       MSLCID = 2067
	SwedishSweden      MSLCID = 1053
	DanishDenmark      MSLCID = 1030
	FinnishFinland     MSLCID = 1035
	NorwegianBokmal    MSLCID = 1044
	GreekGreece        MSLCID = 1032
	TurkishTurkey      MSLCID = 1055

	// África
	AfrikaansSouthAfr MSLCID = 1078
	SwahiliKenya      MSLCID = 1089
)

var LCIDToLanguageTag = map[MSLCID]string{
	EnglishUS:          "en-US",
	EnglishUK:          "en-GB",
	EnglishAustralia:   "en-AU",
	EnglishCanada:      "en-CA",
	EnglishIndia:       "en-IN",
	EnglishIreland:     "en-IE",
	EnglishNZ:          "en-NZ",
	EnglishSouthAfr:    "en-ZA",
	SpanishSpain:       "es-ES",
	SpanishMexico:      "es-MX",
	SpanishArgentina:   "es-AR",
	SpanishColombia:    "es-CO",
	SpanishChile:       "es-CL",
	SpanishPeru:        "es-PE",
	PortugueseBrazil:   "pt-BR",
	PortuguesePortugal: "pt-PT",
	FrenchFrance:       "fr-FR",
	FrenchCanada:       "fr-CA",
	FrenchSwitzerland:  "fr-CH",
	FrenchBelgium:      "fr-BE",
	GermanGermany:      "de-DE",
	GermanAustria:      "de-AT",
	GermanSwitzerland:  "de-CH",
	ChineseChina:       "zh-CN",
	ChineseTaiwan:      "zh-TW",
	ChineseHongKong:    "zh-HK",
	JapaneseJapan:      "ja-JP",
	KoreanKorea:        "ko-KR",
	HindiIndia:         "hi-IN",
	BengaliIndia:       "bn-IN",
	TamilIndia:         "ta-IN",
	Indonesian:         "id-ID",
	MalayMalaysia:      "ms-MY",
	Vietnamese:         "vi-VN",
	ThaiThailand:       "th-TH",
	TagalogPhils:       "tl-PH",
	ArabicSaudiArab:    "ar-SA",
	ArabicEgypt:        "ar-EG",
	ArabicUAE:          "ar-AE",
	HebrewIsrael:       "he-IL",
	PersianIran:        "fa-IR",
	UrduPakistan:       "ur-PK",
	RussianRussia:      "ru-RU",
	Ukrainian:          "uk-UA",
	PolishPoland:       "pl-PL",
	CzechCzechRep:      "cs-CZ",
	SlovakSlovakia:     "sk-SK",
	HungarianHungary:   "hu-HU",
	RomanianRomania:    "ro-RO",
	Bulgarian:          "bg-BG",
	Croatian:           "hr-HR",
	ItalianItaly:       "it-IT",
	ItalianSwitzerland: "it-CH",
	DutchNetherlands:   "nl-NL",
	DutchBelgium:       "nl-BE",
	SwedishSweden:      "sv-SE",
	DanishDenmark:      "da-DK",
	FinnishFinland:     "fi-FI",
	NorwegianBokmal:    "nb-NO",
	GreekGreece:        "el-GR",
	TurkishTurkey:      "tr-TR",
	AfrikaansSouthAfr:  "af-ZA",
	SwahiliKenya:       "sw-KE",
}

// Go Stringformated
func (l MSLCID) String() string {
	if tag, ok := LCIDToLanguageTag[l]; ok {
		return tag
	}
	return "en-US" // Fallback
}

// GetGoLanguageFromMSLCID returns the Go language tag for a given MSLCID
func GetGoLanguageFromMSLCID(l MSLCID) string {
	if tag, ok := LCIDToLanguageTag[l]; ok {
		return tag
	}
	return "en-US" // Fallback
}
