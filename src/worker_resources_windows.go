//go:build windows

package main

import (
	"syscall"
	"unsafe"
)

type workerMemoryStatus struct {
	length, load                         uint32
	totalPhys, availPhys                 uint64
	totalPageFile, availPageFile         uint64
	totalVirtual, availVirtual, extended uint64
}

var workerMemoryAPI = syscall.NewLazyDLL("kernel32.dll").NewProc("GlobalMemoryStatusEx")

func availableMemoryBytes() (uint64, bool) {
	var status workerMemoryStatus
	status.length = uint32(unsafe.Sizeof(status))
	ok, _, _ := workerMemoryAPI.Call(uintptr(unsafe.Pointer(&status)))
	return status.availPhys, ok != 0
}
