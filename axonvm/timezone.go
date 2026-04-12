package axonvm

import (
	"strings"
	"time"
)

// ResolveTimezoneLocation resolves a configured timezone using Go's native time package.
// It returns UTC when the input is empty.
func ResolveTimezoneLocation(name string) (*time.Location, error) {
	tz := strings.TrimSpace(name)
	if tz == "" {
		return time.UTC, nil
	}
	return time.LoadLocation(tz)
}
