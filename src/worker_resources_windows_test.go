//go:build windows

package main

import (
	"testing"
	"unsafe"
)

func TestWindowsMemoryStatus(t *testing.T) {
	if unsafe.Sizeof(workerMemoryStatus{}) != 64 {
		t.Fatal("incorrect MEMORYSTATUSEX layout")
	}
	if available, ok := availableMemoryBytes(); !ok || available == 0 {
		t.Fatal("Windows available-memory query failed")
	}
}
