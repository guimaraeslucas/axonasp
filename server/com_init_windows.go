//go:build windows

package server

import (
	"errors"
	"strings"

	"github.com/go-ole/go-ole"
)

const rpcEChangedMode uint32 = 0x80010106

// comInitialize initializes COM for the current thread.
// Returns (initialized=true) when this call performed COM initialization
// and must be paired with CoUninitialize by the caller.
func comInitialize() (bool, error) {
	err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	if err == nil {
		return true, nil
	} else if isRPCChangedMode(err) {
		return false, nil
	}

	err = ole.CoInitialize(0)
	if err == nil {
		return true, nil
	} else if isRPCChangedMode(err) {
		return false, nil
	}
	return false, err
}

func isRPCChangedMode(err error) bool {
	var oleErr *ole.OleError
	if errors.As(err, &oleErr) {
		if uint32(oleErr.Code()) == rpcEChangedMode {
			return true
		}
	}
	msg := strings.ToLower(err.Error())
	return strings.Contains(msg, "changed mode")
}
