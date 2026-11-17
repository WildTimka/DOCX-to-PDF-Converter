//go:build !windows

package main

import (
	"os"
	"syscall"
	"time"
)

type trackedWord struct{ p *os.Process }

func openTrackedWord(pid int) (*trackedWord, error) {
	p, err := os.FindProcess(pid)
	if err != nil {
		return nil, err
	}
	return &trackedWord{p: p}, nil
}
func (p *trackedWord) wait(timeout time.Duration) bool {
	if p == nil || p.p == nil {
		return true
	}
	deadline := time.Now().Add(timeout)
	for {
		if err := p.p.Signal(syscall.Signal(0)); err != nil {
			return true
		}
		if timeout <= 0 || time.Now().After(deadline) {
			return false
		}
		time.Sleep(25 * time.Millisecond)
	}
}
func (p *trackedWord) kill() error {
	if p == nil || p.p == nil {
		return nil
	}
	return p.p.Kill()
}
func (p *trackedWord) close() {
	if p != nil && p.p != nil {
		_ = p.p.Release()
		p.p = nil
	}
}
