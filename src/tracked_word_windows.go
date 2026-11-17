//go:build windows

package main

import (
	"syscall"
	"time"
)

// trackedWord owns a kernel handle to the exact Word process created by the
// worker. Unlike reopening a numeric PID at shutdown, the handle remains tied
// to that process even if Windows reuses its PID.
type trackedWord struct{ handle syscall.Handle }

func openTrackedWord(pid int) (*trackedWord, error) {
	h, err := syscall.OpenProcess(syscall.PROCESS_TERMINATE|syscall.SYNCHRONIZE, false, uint32(pid))
	if err != nil {
		return nil, err
	}
	return &trackedWord{handle: h}, nil
}

func (p *trackedWord) wait(timeout time.Duration) bool {
	if p == nil || p.handle == 0 {
		return true
	}
	ms := uint32(0)
	if timeout > 0 {
		ms = uint32(timeout / time.Millisecond)
	}
	status, err := syscall.WaitForSingleObject(p.handle, ms)
	return err == nil && status == syscall.WAIT_OBJECT_0
}

func (p *trackedWord) kill() error {
	if p == nil || p.handle == 0 {
		return nil
	}
	return syscall.TerminateProcess(p.handle, 1)
}

func (p *trackedWord) close() {
	if p != nil && p.handle != 0 {
		_ = syscall.CloseHandle(p.handle)
		p.handle = 0
	}
}
