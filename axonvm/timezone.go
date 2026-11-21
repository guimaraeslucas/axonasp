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
	_ "time/tzdata"
)

// ResolveTimezoneLocation resolves a configured timezone using Go's native time package.
// It returns UTC when the input is empty.
func ResolveTimezoneLocation(name string) (*time.Location, error) {
	tz := strings.TrimSpace(name)
	if tz == "" {
		return time.UTC, nil
	}
	if loc, ok := parseFixedZone(tz); ok {
		return loc, nil
	}
	return time.LoadLocation(tz)
}

func parseFixedZone(tz string) (*time.Location, bool) {
	var prefix string
	if strings.HasPrefix(tz, "UTC") {
		prefix = "UTC"
	} else if strings.HasPrefix(tz, "GMT") {
		prefix = "GMT"
	} else {
		return nil, false
	}

	offsetStr := strings.TrimSpace(tz[len(prefix):])
	if offsetStr == "" {
		return time.UTC, true
	}

	sign := 1
	switch offsetStr[0] {
	case '+':
		sign = 1
		offsetStr = offsetStr[1:]
	case '-':
		sign = -1
		offsetStr = offsetStr[1:]
	default:
		return nil, false
	}

	var hours, minutes int
	var err error
	if strings.Contains(offsetStr, ":") {
		parts := strings.Split(offsetStr, ":")
		if len(parts) == 2 {
			hours, err = strconv.Atoi(parts[0])
			if err == nil {
				minutes, err = strconv.Atoi(parts[1])
			}
		} else {
			return nil, false
		}
	} else if len(offsetStr) == 4 {
		hours, err = strconv.Atoi(offsetStr[0:2])
		if err == nil {
			minutes, err = strconv.Atoi(offsetStr[2:4])
		}
	} else {
		hours, err = strconv.Atoi(offsetStr)
	}

	if err != nil {
		return nil, false
	}

	totalSeconds := sign * (hours*3600 + minutes*60)
	return time.FixedZone(tz, totalSeconds), true
}

// normalizeSQLDateTime rebinds timezone-naive SQL date/time values decoded in UTC
// to the VM runtime location while preserving timezone-aware types such as DATETIMEOFFSET.
func normalizeSQLDateTime(databaseType string, v any, loc *time.Location) any {
	if v == nil {
		return nil
	}
	t, ok := v.(time.Time)
	if !ok || t.IsZero() {
		return v
	}
	if loc == nil {
		loc = time.UTC
	}

	normalized := strings.ToUpper(strings.TrimSpace(databaseType))
	if parenIdx := strings.IndexByte(normalized, '('); parenIdx >= 0 {
		normalized = strings.TrimSpace(normalized[:parenIdx])
	}

	switch normalized {
	case "DATETIMEOFFSET", "TIMESTAMPTZ", "TIMESTAMP WITH TIME ZONE", "TIMESTAMP WITH LOCAL TIME ZONE", "TIMETZ", "TIME WITH TIME ZONE":
		return t
	case "TIME", "TIME WITHOUT TIME ZONE":
		return time.Date(1899, time.December, 30, t.Hour(), t.Minute(), t.Second(), t.Nanosecond(), loc)
	case "DATETIME", "SMALLDATETIME", "DATETIME2", "DATE", "TIMESTAMP", "TIMESTAMP WITHOUT TIME ZONE":
		_, offset := t.Zone()
		if offset == 0 && (t.Location() == time.UTC || t.Location() == nil || t.Location().String() == "") {
			return time.Date(t.Year(), t.Month(), t.Day(), t.Hour(), t.Minute(), t.Second(), t.Nanosecond(), loc)
		}
		return t
	case "SQL_VARIANT", "":
		if t.Location() == time.UTC {
			if t.Year() == 1 && t.Month() == time.January && t.Day() == 1 && (t.Hour() != 0 || t.Minute() != 0 || t.Second() != 0 || t.Nanosecond() != 0) {
				return time.Date(1899, time.December, 30, t.Hour(), t.Minute(), t.Second(), t.Nanosecond(), loc)
			}
			return time.Date(t.Year(), t.Month(), t.Day(), t.Hour(), t.Minute(), t.Second(), t.Nanosecond(), loc)
		}
		return t
	default:
		return t
	}
}
